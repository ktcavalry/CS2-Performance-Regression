# -*- coding: utf-8 -*-
"""
steam_cs_scraper_final_limited.py
功能: Steam CS2 战绩抓取 + CSV 保存 + KD/Score 趋势图 + 实时更新 + 抓取失败截图 + 只抓指定玩家 + 保留最新 N 条记录
"""
import os
import time
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ---------------- 配置 ----------------
CSV_FILE = "steam_cs_stats.csv"
TREND_FILE = "kd_score_trend.png"
BROWSER = "chrome"  # chrome / edge
HEADLESS = False     # True=无头模式，False=显示浏览器
FETCH_INTERVAL = 300
MAX_MATCHES = 50  # 保留最新比赛条数

TARGET_PLAYER = "月生夜夜"  # 玩家昵称
TARGET_STEAMID = "76561199764696725"  # 玩家 SteamID64
MATCH_HISTORY_TAB = "matchhistorypremier"  # 可改为 matchhistory 或 matchhistorypremier

SCREENSHOT_DIR = "screenshots"  # 抓取失败截图保存目录

SMTP_SERVER = ""
SMTP_PORT = 465
SMTP_USER = ""
SMTP_PASS = ""
EMAIL_TO = []

# ---------------- 工具函数 ----------------
def parse_int(text):
    try:
        return int(str(text).strip().replace("★","").replace(",",""))
    except:
        return 0

def parse_float(text):
    try:
        return float(str(text).strip().replace("%","").replace(",",""))
    except:
        return 0.0

def parse_mvp(text):
    try:
        return int(str(text).strip().replace("★",""))
    except:
        return 0

def parse_steam_time(time_str):
    try:
        time_str = time_str.strip()
        # 移除末尾的 GMT 或其他时区信息
        if " GMT" in time_str:
            time_str = time_str.replace(" GMT", "").strip()
        
        # 解析GMT时间，然后加8小时转换为北京时间（UTC+8）
        gmt_time = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
        beijing_time = gmt_time + timedelta(hours=8)
        
        return beijing_time.strftime("%Y-%m-%d %H:%M:%S")
    except:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# 胜负判定函数
def get_match_result(table, player_name, target_player):
    """
    从表格判定胜负
    返回: (我队分数, 敌队分数, 胜负标记)
    """
    try:
        # 查找分数行 - 分数在 <td class="csgo_scoreboard_score">
        score_cell = table.find_element(By.CSS_SELECTOR, "td.csgo_scoreboard_score")
        score_text = score_cell.text.strip()  # 例: "8 : 13"
        scores = [int(x.strip()) for x in score_text.split(":")]
        if len(scores) != 2:
            return 0, 0, "Unknown"
        
        left_score, right_score = scores[0], scores[1]
        
        # 判断玩家在上方还是下方
        # 查找所有玩家行（分数行之前的都是一队，之后的是另一队）
        all_rows = table.find_elements(By.TAG_NAME, "tr")
        score_row_index = -1
        
        # 找到分数行的索引
        for i, row in enumerate(all_rows):
            try:
                row.find_element(By.CSS_SELECTOR, "td.csgo_scoreboard_score")
                score_row_index = i
                break
            except:
                pass
        
        if score_row_index == -1:
            return left_score, right_score, "Unknown"
        
        # 检查玩家是否在上方（分数行之前）
        player_in_top = False
        for i in range(score_row_index):
            try:
                nickname_elem = all_rows[i].find_element(By.CSS_SELECTOR, "div.playerNickname")
                if target_player in nickname_elem.text:
                    player_in_top = True
                    break
            except:
                pass
        
        # 判断胜负
        if player_in_top:
            team_score = left_score
            enemy_score = right_score
        else:
            team_score = right_score
            enemy_score = left_score
        
        if team_score > enemy_score:
            result = "Win"
        elif team_score < enemy_score:
            result = "Loss"
        else:
            result = "Draw"
        
        return team_score, enemy_score, result
    except Exception as e:
        print(f"  ⚠️ 判定胜负失败: {e}")
        return 0, 0, "Unknown"

# ---------------- 浏览器初始化 ----------------
def init_driver(browser=BROWSER, headless=HEADLESS):
    if browser.lower() == "chrome":
        opts = ChromeOptions()
        if headless:
            opts.add_argument("--headless=new")
        opts.add_argument("--start-maximized")
        service = ChromeService(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=opts)
    elif browser.lower() == "edge":
        opts = EdgeOptions()
        if headless:
            opts.add_argument("--headless=new")
        opts.add_argument("--start-maximized")
        service = EdgeService(EdgeChromiumDriverManager().install())
        return webdriver.Edge(service=service, options=opts)
    else:
        raise ValueError("BROWSER must be 'chrome' or 'edge'")

# ---------------- CSV 管理 ----------------
def load_csv():
    if os.path.exists(CSV_FILE):
        return pd.read_csv(CSV_FILE, encoding='utf-8-sig')
    df = pd.DataFrame(columns=["MatchID","DateTime","Mode","Player Name","Ping","K","A","D","MVP","HSP","Score","TeamScore","EnemyScore","Result"])
    df.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')
    print(f"✅ 已创建空 CSV 文件: {CSV_FILE}")
    return df

def save_csv(df):
    # 保存所有记录
    df.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')

# ---------------- 点击加载更多 ----------------
def load_more(driver):
    load_count = 0
    consecutive_failures = 0
    max_consecutive_failures = 3  # 连续3次失败才停止
    
    while consecutive_failures < max_consecutive_failures:
        try:
            btn = driver.find_element(By.ID, "load_more_button")
            if btn.is_displayed() and btn.is_enabled():
                driver.execute_script("arguments[0].click();", btn)
                load_count += 1
                print(f"🔹 第 {load_count} 次加载更多历史记录")
                consecutive_failures = 0  # 重置失败计数
                time.sleep(2)  # 增加等待时间，让页面充分加载
            else:
                consecutive_failures += 1
                break
        except Exception as e:
            consecutive_failures += 1
            print(f"⚠️ 加载失败 (第{consecutive_failures}次): {e}")
            time.sleep(1)
    
    if load_count > 0:
        print(f"✅ 共加载 {load_count} 次，总计已加载全部可用数据")
    else:
        print("⚠️ 未找到加载更多按钮或已加载全部数据")

# ---------------- 等待比赛表格加载 ----------------
def wait_for_matches(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.csgo_scoreboard_inner_right"))
        )
        print("✅ 比赛表格已加载")
        return True
    except Exception as e:
        print(f"⚠️ 等待表格超时: {e}")
        # 尝试保存页面用于调试
        screenshot_path = os.path.join(SCREENSHOT_DIR, "page_load_failed.png")
        os.makedirs(SCREENSHOT_DIR, exist_ok=True)
        driver.save_screenshot(screenshot_path)
        print(f"📷 页面截图已保存: {screenshot_path}")
        return False

# ---------------- 抓取比赛 ----------------
def scrape_all(driver):
    if not wait_for_matches(driver):
        print("⚠️ 未检测到比赛表格")
        return []
    load_more(driver)
    left_tables = driver.find_elements(By.CSS_SELECTOR, "table.csgo_scoreboard_inner_left")
    right_tables = driver.find_elements(By.CSS_SELECTOR, "table.csgo_scoreboard_inner_right")
    rows_out = []

    for idx, (ltbl, rtbl) in enumerate(zip(left_tables, right_tables)):
        try:
            mode = ltbl.find_element(By.TAG_NAME,"td").text.strip()
            dt_text = ltbl.find_elements(By.TAG_NAME,"td")[1].text.strip()
            match_time = parse_steam_time(dt_text)
            match_id = match_time  # 使用比赛时间作为MatchID
            print(f"🟢 正在抓取比赛: {mode} @ {match_time}")

            trs = rtbl.find_elements(By.TAG_NAME,"tr")[1:]
            for tr in trs:
                player_name = "Unknown"
                try:
                    # 尝试选择器：div.playerNickname a (查找链接)
                    player_name = tr.find_element(By.CSS_SELECTOR,"div.playerNickname a").text.strip()
                except:
                    try:
                        # 备选选择器 2：.linkTitle (直接查找链接)
                        player_name = tr.find_element(By.CSS_SELECTOR,".playerNickname a.linkTitle").text.strip()
                    except:
                        try:
                            # 备选选择器 3：按类名
                            elem = tr.find_element(By.CLASS_NAME,"playerNickname")
                            player_name = elem.find_element(By.TAG_NAME,"a").text.strip()
                        except:
                            pass
                if player_name != TARGET_PLAYER:
                    continue
                cells = tr.find_elements(By.TAG_NAME,"td")
                if len(cells)<8:
                    continue
                ping = parse_int(cells[1].text)
                k = parse_int(cells[2].text)
                a = parse_int(cells[3].text)
                d = parse_int(cells[4].text)
                mvp = parse_mvp(cells[5].text)
                hsp = parse_float(cells[6].text)
                score = parse_int(cells[7].text)
                
                # 判断胜负
                team_score, enemy_score, result = get_match_result(rtbl, player_name, TARGET_PLAYER)
                
                print(f"  - 玩家: {player_name}, K/D/A: {k}/{d}/{a}, MVP: {mvp}, Score: {score}, Ping: {ping}, 比分: {team_score}:{enemy_score}, 结果: {result}")
                rows_out.append([match_id, match_time, mode, player_name, ping, k, a, d, mvp, hsp, score, team_score, enemy_score, result])
        except Exception as e:
            print(f"⚠️ 某场比赛抓取失败: {e}")
            screenshot_path = os.path.join(SCREENSHOT_DIR, f"fail_scrape_{idx}.png")
            os.makedirs(SCREENSHOT_DIR, exist_ok=True)
            driver.save_screenshot(screenshot_path)
            print(f"📷 已保存抓取失败截图: {screenshot_path}")
            continue
    return rows_out

# ---------------- 实时抓取最新比赛 ----------------
def scrape_latest(driver, df_all):
    if not wait_for_matches(driver):
        return False, df_all
    load_more(driver)
    left_tables = driver.find_elements(By.CSS_SELECTOR, "table.csgo_scoreboard_inner_left")
    right_tables = driver.find_elements(By.CSS_SELECTOR, "table.csgo_scoreboard_inner_right")
    for idx, (ltbl, rtbl) in enumerate(zip(left_tables, right_tables)):
        try:
            mode = ltbl.find_element(By.TAG_NAME,"td").text.strip()
            dt_text = ltbl.find_elements(By.TAG_NAME,"td")[1].text.strip()
            match_time = parse_steam_time(dt_text)
            match_id = match_time  # 使用比赛时间作为MatchID
            print(f"🟢 实时抓取比赛: {mode} @ {match_time}")

            trs = rtbl.find_elements(By.TAG_NAME,"tr")[1:]
            match_rows = []
            for tr in trs:
                player_name = "Unknown"
                try:
                    # 尝试选择器：div.playerNickname a
                    player_name = tr.find_element(By.CSS_SELECTOR,"div.playerNickname a").text.strip()
                except:
                    try:
                        # 备选选择器 2：.linkTitle
                        player_name = tr.find_element(By.CSS_SELECTOR,".playerNickname a.linkTitle").text.strip()
                    except:
                        try:
                            # 备选选择器 3：按类名
                            elem = tr.find_element(By.CLASS_NAME,"playerNickname")
                            player_name = elem.find_element(By.TAG_NAME,"a").text.strip()
                        except:
                            pass
                if player_name != TARGET_PLAYER:
                    continue
                cells = tr.find_elements(By.TAG_NAME,"td")
                if len(cells)<8:
                    continue
                ping = parse_int(cells[1].text)
                k = parse_int(cells[2].text)
                a = parse_int(cells[3].text)
                d = parse_int(cells[4].text)
                mvp = parse_mvp(cells[5].text)
                hsp = parse_float(cells[6].text)
                score = parse_int(cells[7].text)
                
                # 判断胜负
                team_score, enemy_score, result = get_match_result(rtbl, player_name, TARGET_PLAYER)
                
                print(f"  - 玩家: {player_name}, K/D/A: {k}/{d}/{a}, MVP: {mvp}, Score: {score}, Ping: {ping}, 比分: {team_score}:{enemy_score}, 结果: {result}")
                match_rows.append([match_id, match_time, mode, player_name, ping, k, a, d, mvp, hsp, score, team_score, enemy_score, result])

            if not match_rows:
                continue
            df_new = pd.DataFrame(match_rows, columns=df_all.columns)
            df_all = pd.concat([df_all, df_new], ignore_index=True)
            for col in ["Ping","K","A","D","MVP","Score","HSP","TeamScore","EnemyScore"]:
                if col in df_all.columns:
                    if col=="HSP":
                        df_all[col] = pd.to_numeric(df_all[col], errors='coerce').fillna(0)
                    else:
                        df_all[col] = pd.to_numeric(df_all[col], errors='coerce').fillna(0).astype(int)
            df_all.drop_duplicates(subset=["MatchID","Player Name","DateTime"], inplace=True)
            save_csv(df_all)
            return True, df_all
        except Exception as e:
            print(f"⚠️ 实时抓取失败: {e}")
            screenshot_path = os.path.join(SCREENSHOT_DIR, f"fail_latest_{idx}.png")
            os.makedirs(SCREENSHOT_DIR, exist_ok=True)
            driver.save_screenshot(screenshot_path)
            print(f"📷 已保存抓取失败截图: {screenshot_path}")
            continue
    return False, df_all

# ---------------- 绘图 ----------------
def update_trend(df):
    if df.empty:
        return
    
    # KD 保留两位小数
    df["KD"] = df.apply(lambda x: round(x["K"]/x["D"],2) if x["D"]>0 else round(x["K"],2), axis=1)
    df_sorted = df.sort_values("DateTime").tail(50).reset_index(drop=True)  # 只显示最新50场
    
    # 设置中文字体
    plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False
    
    # 创建图表
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10))
    fig.suptitle(f"{TARGET_PLAYER} CS2 战绩趋势分析", fontsize=18, fontweight='bold', y=0.995)
    
    # --- 第一个图：KD趋势 ---
    ax1.plot(df_sorted.index, df_sorted["KD"], marker='o', linewidth=2.5, markersize=8, 
             color='#2E86AB', label="KD值", zorder=2)
    ax1.fill_between(df_sorted.index, df_sorted["KD"], alpha=0.2, color='#2E86AB')
    
    # 标注数据点
    for i, (idx, kd) in enumerate(zip(df_sorted.index, df_sorted["KD"])):
        ax1.text(idx, kd + 0.05, f'{kd:.2f}', ha='center', va='bottom', fontsize=9)
    
    ax1.set_xlabel("比赛序号", fontsize=12, fontweight='bold')
    ax1.set_ylabel("KD值", fontsize=12, fontweight='bold')
    ax1.set_title("KD值趋势 (击杀/死亡比)", fontsize=13, fontweight='bold', pad=10)
    ax1.grid(True, alpha=0.3, linestyle='--')
    ax1.legend(loc='upper left', fontsize=11)
    ax1.set_facecolor('#F8F9FA')
    
    # --- 第二个图：分数趋势（含异常值处理）---
    # 计算分数的四分位数，识别异常值
    scores = df_sorted["Score"]
    Q1 = scores.quantile(0.25)
    Q3 = scores.quantile(0.75)
    IQR = Q3 - Q1
    threshold = Q3 + 1.5 * IQR
    max_normal = min(threshold, scores.quantile(0.95))  # 设置显示上限
    
    # 绘制柱子（异常值截断）
    bar_heights = [min(score, max_normal) for score in scores]
    bars = ax2.bar(df_sorted.index, bar_heights, width=0.7, color='#A23B72', alpha=0.8, label="比赛分数")
    
    # 为异常值添加破折号标记
    for idx, (i, score) in enumerate(zip(df_sorted.index, scores)):
        if score > max_normal:
            ax2.plot([i - 0.35, i + 0.35], [max_normal, max_normal], 'r--', linewidth=2, zorder=10)
    
    # 绘制分数走势线（也要截断）
    line_heights = [min(score, max_normal) for score in scores]
    ax2.plot(df_sorted.index, line_heights, marker='s', linewidth=2, markersize=6, 
             color='#F18F01', label="分数走势", zorder=3)
    
    # 标注分数（显示完整值）
    for i, (idx, score) in enumerate(zip(df_sorted.index, scores)):
        if score > max_normal:
            # 异常值用红色显示并标注完整值
            ax2.text(idx, max_normal + 2, f'{int(score)}', ha='center', va='bottom', 
                    fontsize=9, color='red', fontweight='bold')
        else:
            ax2.text(idx, score + 1, str(int(score)), ha='center', va='bottom', fontsize=9)
    
    ax2.set_xlabel("比赛序号", fontsize=12, fontweight='bold')
    ax2.set_ylabel("比赛分数", fontsize=12, fontweight='bold')
    ax2.set_title("每场比赛分数", fontsize=13, fontweight='bold', pad=10)
    ax2.grid(True, alpha=0.3, linestyle='--', axis='y')
    ax2.legend(loc='upper left', fontsize=11)
    ax2.set_facecolor('#F8F9FA')
    ax2.set_ylim(0, max_normal * 1.2)
    
    plt.tight_layout()
    plt.savefig(TREND_FILE, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"📊 趋势图已更新: {TREND_FILE}")

# ---------------- 统计与导出 ----------------
def generate_statistics(df):
    """生成详细的数据统计"""
    if df.empty:
        print("⚠️ 无数据可统计")
        return
    
    print("\n" + "="*60)
    print(f"📈 玩家 {TARGET_PLAYER} 战绩统计".center(60))
    print("="*60)
    
    # 获取时间范围
    df_time = df.copy()
    df_time['DateTime'] = pd.to_datetime(df_time['DateTime'], errors='coerce')
    first_time = df_time['DateTime'].min()
    last_time = df_time['DateTime'].max()
    days_span = (last_time - first_time).days
    
    total_matches = len(df)
    total_kills = df["K"].sum()
    total_deaths = df["D"].sum()
    total_assists = df["A"].sum()
    total_score = df["Score"].sum()
    total_mvp = df["MVP"].sum()
    
    kd_ratio = round(total_kills / total_deaths, 2) if total_deaths > 0 else total_kills
    avg_kills = round(total_kills / total_matches, 2)
    avg_deaths = round(total_deaths / total_matches, 2)
    avg_score = round(total_score / total_matches, 2)
    avg_hsp = df["HSP"].mean()
    
    print(f"\n📅 时间范围：")
    print(f"  • 首场比赛: {first_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  • 最后比赛: {last_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  • 统计周期: {days_span} 天")
    
    print(f"\n📊 基础统计：")
    print(f"  • 总比赛数: {total_matches} 场")
    print(f"  • 总击杀: {total_kills} 次")
    print(f"  • 总死亡: {total_deaths} 次")
    print(f"  • 总助攻: {total_assists} 次")
    print(f"  • MVP次数: {total_mvp} 次")
    print(f"  • 总分数: {total_score} 分")
    
    print(f"\n📈 平均数据：")
    print(f"  • 场均击杀: {avg_kills}")
    print(f"  • 场均死亡: {avg_deaths}")
    print(f"  • 场均分数: {avg_score}")
    print(f"  • 平均HSP: {avg_hsp:.1f}%")
    
    print(f"\n🎯 总体表现：")
    print(f"  • 总KD值: {kd_ratio}")
    print(f"  • MVP率: {round(total_mvp/total_matches*100, 1)}%")
    
    # 胜率统计
    if "Result" in df.columns:
        wins = len(df[df["Result"] == "Win"])
        losses = len(df[df["Result"] == "Loss"])
        draws = len(df[df["Result"] == "Draw"])
        win_rate = round(wins / total_matches * 100, 1) if total_matches > 0 else 0
        print(f"  • 胜率: {win_rate}%")
    else:
        wins = losses = draws = 0
        win_rate = 0
    
    print(f"\n🏆 胜负统计：")
    print(f"  • 胜场: {wins}")
    print(f"  • 负场: {losses}")
    print(f"  • 平局: {draws}")
    
    # 按模式统计
    print(f"\n🗺️  按模式分类：")
    mode_stats = df.groupby("Mode").agg({
        "MatchID": "count",
        "K": "sum",
        "D": "sum",
        "Score": "sum"
    }).rename(columns={"MatchID": "场数"})
    for mode, row in mode_stats.iterrows():
        kd = round(row["K"] / row["D"], 2) if row["D"] > 0 else row["K"]
        print(f"  • {mode}: {int(row['场数'])}场 | K:{int(row['K'])} D:{int(row['D'])} | KD:{kd} | 总分:{int(row['Score'])}")
    
    print("\n" + "="*60 + "\n")

def export_all_data(df):
    """导出完整数据到多个Excel表"""
    if df.empty:
        print("⚠️ 无数据可导出")
        return
    
    from openpyxl.styles import Font, Alignment
    excel_file = "steam_cs_stats_完整数据.xlsx"
    
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        # 表1: 所有原始数据
        df_export = df.copy()
        df_export.to_excel(writer, sheet_name='原始数据', index=False)
        ws1 = writer.sheets['原始数据']
        for column in ws1.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws1.column_dimensions[column_letter].width = adjusted_width
            for cell in column:
                cell.font = Font(size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 表2: 统计汇总
        total_matches = len(df)
        wins = len(df[df["Result"] == "Win"]) if "Result" in df.columns else 0
        losses = len(df[df["Result"] == "Loss"]) if "Result" in df.columns else 0
        draws = len(df[df["Result"] == "Draw"]) if "Result" in df.columns else 0
        win_rate = round(wins / total_matches * 100, 1) if total_matches > 0 else 0
        
        stats_data = {
            '统计项': ['总比赛数', '胜场', '负场', '平局', '胜率(%)', '总击杀', '总死亡', '总助攻', 'MVP次数', '总分数'],
            '数值': [
                total_matches,
                wins,
                losses,
                draws,
                win_rate,
                df["K"].sum(),
                df["D"].sum(),
                df["A"].sum(),
                df["MVP"].sum(),
                df["Score"].sum()
            ]
        }
        pd.DataFrame(stats_data).to_excel(writer, sheet_name='统计汇总', index=False)
        ws2 = writer.sheets['统计汇总']
        for column in ws2.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws2.column_dimensions[column_letter].width = adjusted_width
            for cell in column:
                cell.font = Font(size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 表3: 模式分析
        mode_analysis = df.groupby("Mode").agg({
            "MatchID": "count",
            "K": ["sum", "mean"],
            "D": ["sum", "mean"],
            "Score": ["sum", "mean"],
            "MVP": "sum"
        }).round(2)
        mode_analysis.columns = ['场数', '总击杀', '场均击杀', '总死亡', '场均死亡', '总分数', '场均分数', 'MVP次数']
        mode_analysis.to_excel(writer, sheet_name='模式分析')
        ws3 = writer.sheets['模式分析']
        for column in ws3.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws3.column_dimensions[column_letter].width = adjusted_width
            for cell in column:
                cell.font = Font(size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 表4: 每场比赛详情
        df_detail = df.copy()
        df_detail['KD比'] = (df_detail['K'] / df_detail['D']).round(2)
        df_detail = df_detail[['DateTime', 'Mode', 'K', 'D', 'A', 'KD比', 'MVP', 'HSP', 'Score', 'TeamScore', 'EnemyScore', 'Result']]
        df_detail.to_excel(writer, sheet_name='比赛详情', index=False)
        ws4 = writer.sheets['比赛详情']
        for column in ws4.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws4.column_dimensions[column_letter].width = adjusted_width
            for cell in column:
                cell.font = Font(size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
    
    print(f"✅ 完整数据已导出: {excel_file}")


# ---------------- 邮件通知 ----------------
def send_email(subject, content):
    if not SMTP_SERVER or not SMTP_USER or not SMTP_PASS or not EMAIL_TO:
        return
    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = ",".join(EMAIL_TO)
    msg["Subject"] = subject
    msg.attach(MIMEText(content,"plain"))
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as s:
            s.login(SMTP_USER, SMTP_PASS)
            s.sendmail(SMTP_USER, EMAIL_TO, msg.as_string())
    except:
        pass

# ---------------- 主程序 ----------------
def main():
    driver = init_driver()
    steam_url = f"https://steamcommunity.com/profiles/{TARGET_STEAMID}/gcpd/730/?tab={MATCH_HISTORY_TAB}"
    print(f"正在打开玩家 {TARGET_PLAYER} 的战绩页面: {steam_url}")
    driver.get(steam_url)
    print("请确保已登录 Steam 并等待页面加载完成，等待中…")
    time.sleep(8)  # 增加等待时间至8秒
    print("✓ 页面加载完成，开始处理…\n")

    df_all = load_csv()
    # 判断是否是第一次运行（CSV为空或只有表头）
    is_first_run = len(df_all) == 0
    
    if is_first_run:
        print("\n🆕 检测到首次运行，将自动爬取全部历史数据...")
        mode = "1"
    else:
        mode = input("\n选择模式：1=一次性抓取全部历史  2=实时监控（默认2） 输入1或2：").strip() or "2"
    try:
        if mode=="1":
            rows = scrape_all(driver)
            if rows:
                df_new = pd.DataFrame(rows, columns=df_all.columns)
                df_all = pd.concat([df_all, df_new], ignore_index=True)
                for col in ["Ping","K","A","D","MVP","Score","HSP"]:
                    if col in df_all.columns:
                        if col=="HSP":
                            df_all[col] = pd.to_numeric(df_all[col], errors='coerce').fillna(0)
                        else:
                            df_all[col] = pd.to_numeric(df_all[col], errors='coerce').fillna(0).astype(int)
                df_all.drop_duplicates(subset=["MatchID","Player Name","DateTime"], inplace=True)
                save_csv(df_all)
                update_trend(df_all)
                generate_statistics(df_all)
                export_all_data(df_all)
                print(f"✅ 已抓取并保存 {len(rows)} 条记录。")
            else:
                print("⚠️ 未抓取到记录，请检查登录状态或页面选择器。")
        else:
            print("进入实时监控模式，按 Ctrl+C 停止")
            while True:
                new, df_all = scrape_latest(driver, df_all)
                if new:
                    update_trend(df_all)
                    generate_statistics(df_all)
                    send_email("CS2 新战绩更新", f"抓取到 {TARGET_PLAYER} 的新战绩并已保存。")
                    print("✅ 发现新战绩并已处理。")
                else:
                    print(f"未发现新战绩，{FETCH_INTERVAL} 秒后重试…")
                time.sleep(FETCH_INTERVAL)
    except KeyboardInterrupt:
        print("已停止。")
    finally:
        try:
            driver.quit()
        except:
            pass

if __name__=="__main__":
    main()
