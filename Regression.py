import pandas as pd
import matplotlib.pyplot as plt

from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_error, r2_score

# =========================
# 1. 读取数据
# =========================
df = pd.read_csv("steam_cs_stats.csv")
df = df.sort_values("DateTime").reset_index(drop=True)

# =========================
# 2. 基础特征工程
# =========================
df["KD"] = df["K"] / df["D"].replace(0, 1)
df["Win"] = df["Result"].map({"Win": 1, "Loss": 0, "Draw": 0.5})

window = 5
df["Score_mean_5"] = df["Score"].rolling(window).mean()
df["HSP_mean_5"] = df["HSP"].rolling(window).mean()
df["MVP_mean_5"] = df["MVP"].rolling(window).mean()
df["Ping_mean_5"] = df["Ping"].rolling(window).mean()
df["WinRate_5"] = df["Win"].rolling(window).mean()

# 趋势特征
df["KD_trend_5"] = df["KD"].rolling(window).apply(
    lambda x: x.iloc[-1] - x.iloc[0]
)

# =========================
# 3. 预测目标：未来10场平均KD
# =========================
df["target_KD_10"] = df["KD"].shift(-10).rolling(10).mean()

data = df.dropna().reset_index(drop=True)

features = [
    "Score_mean_5",
    "HSP_mean_5",
    "MVP_mean_5",
    "Ping_mean_5",
    "WinRate_5",
    "KD_trend_5",
]

X = data[features]
y = data["target_KD_10"]

# 时间序列切分
split = int(len(data) * 0.8)
X_train, X_test = X[:split], X[split:]
y_train, y_test = y[:split], y[split:]

# =========================
# 4. RandomForest 回归
# =========================
rf = RandomForestRegressor(
    n_estimators=300,
    max_depth=6,
    min_samples_leaf=10,
    random_state=42,
    n_jobs=-1
)

rf.fit(X_train, y_train)
y_pred = rf.predict(X_test)

# =========================
# 5. 评估
# =========================
mae = mean_absolute_error(y_test, y_pred)
r2 = r2_score(y_test, y_pred)

print("📊 RandomForest KD_10 预测")
print(f"样本数量: {len(data)}")
print(f"MAE: {mae:.4f}")
print(f"R2: {r2:.4f}")

# =========================
# 6. 特征重要性
# =========================
importances = pd.Series(
    rf.feature_importances_,
    index=features
).sort_values(ascending=False)

print("\n🌲 特征重要性:")
for feat, imp in importances.items():
    print(f"  {feat}: {imp:.4f}")

# =========================
# 7. 可视化
# =========================
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 预测结果图
fig, ax = plt.subplots(figsize=(14, 6))
ax.plot(y_test.values, label="实际未来10场平均KD", marker='o', linewidth=2, markersize=6, alpha=0.7)
ax.plot(y_pred, label="预测未来10场平均KD", marker='s', linewidth=2, markersize=6, alpha=0.7)
ax.legend(fontsize=12)
ax.set_xlabel("测试集样本", fontsize=12)
ax.set_ylabel("KD值", fontsize=12)
ax.set_title("RandomForest 未来10场平均KD预测", fontsize=14, fontweight='bold')
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("prediction_KD_10_RF.png", dpi=150)
print("\n✅ 预测结果图已保存: prediction_KD_10_RF.png")

# 特征重要性图
fig, ax = plt.subplots(figsize=(10, 6))
importances.plot(kind="barh", ax=ax, color='steelblue')
ax.set_xlabel("特征重要性", fontsize=12)
ax.set_ylabel("特征", fontsize=12)
ax.set_title("RandomForest 特征重要性分析", fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig("feature_importance_RF.png", dpi=150)
print("✅ 特征重要性图已保存: feature_importance_RF.png")

plt.show()
