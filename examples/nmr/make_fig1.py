# 図1(共鳴周波数 f vs 磁場 B と最小二乗直線)を figures/fig1_f_vs_B.png に生成する。
#   python make_fig1.py
# データは data/nmr_measurements.csv。レポート用のモノクロ・単系列の散布図。
import csv
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))

# 日本語フォント(環境にあるものを使う)
matplotlib.rcParams["font.family"] = ["Noto Sans CJK JP", "IPAPGothic", "IPAGothic"]
matplotlib.rcParams["axes.unicode_minus"] = False

f_MHz, B_mT = [], []
with open(os.path.join(HERE, "data", "nmr_measurements.csv")) as fp:
    for row in csv.DictReader(r for r in fp if not r.startswith("#")):
        f_MHz.append(float(row["f_MHz"]))
        B_mT.append(float(row["B_mT"]))
f = np.array(f_MHz) * 1e6   # Hz
B = np.array(B_mT) * 1e-3   # T

# 配布資料の式(中心化した最小二乗法)で傾きを求め、切片は平均値を通す条件で決める
dB, df = B - B.mean(), f - f.mean()
slope = np.sum(dB * df) / np.sum(dB * dB)   # [Hz/T]
intercept = f.mean() - slope * B.mean()      # [Hz]

fig, ax = plt.subplots(figsize=(6.0, 4.0), dpi=200)

bx = np.linspace(180, 268, 2)  # mT
ax.plot(bx, (slope * bx * 1e-3 + intercept) / 1e6, color="#444444", lw=1.2, zorder=2,
        label="最小二乗直線")
ax.plot(B_mT, f_MHz, "o", ms=7, mfc="white", mec="black", mew=1.4, zorder=3,
        label="測定値")

ax.annotate(rf"$df/dB = {slope / 1e7:.4f} \times 10^{{7}}$ Hz/T",
            xy=(0.05, 0.86), xycoords="axes fraction", fontsize=10, color="#222222")

ax.set_xlabel("磁場 B [mT]", fontsize=11)
ax.set_ylabel("周波数 f [MHz]", fontsize=11)
ax.set_xlim(180, 268)
ax.set_ylim(7.5, 11.5)
ax.grid(True, color="#dddddd", lw=0.6)
ax.tick_params(labelsize=10)
for s in ("top", "right"):
    ax.spines[s].set_color("#888888")
for s in ("left", "bottom"):
    ax.spines[s].set_color("#444444")
ax.legend(fontsize=10, loc="lower right", framealpha=1.0, edgecolor="#bbbbbb")

fig.tight_layout()
out = os.path.join(HERE, "figures", "fig1_f_vs_B.png")
fig.savefig(out, facecolor="white")
print("saved:", out, f"(slope = {slope:.6e} Hz/T, intercept = {intercept:.4e} Hz)")
