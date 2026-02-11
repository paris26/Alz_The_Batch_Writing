"""
Generate 4 presentation charts for the AI Alzheimer Thesis Presentation.
All charts: 300 DPI, transparent background, presentation color palette, serif font.
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import os

# ── Paths ──
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generated_charts")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Color Palette ──
DARK_BG    = "#0D1117"
LIGHT_BG   = "#F7F5F0"
COPPER     = "#C17F3A"
BLUE       = "#3B82B6"
RED        = "#DC4A4A"
GREEN      = "#5B8C6B"
TEXT_DARK  = "#2D2D2D"
TEXT_LIGHT = "#E8E4DE"
CITE_GRAY  = "#8B8680"

# ── Font setup ──
plt.rcParams.update({
    'font.family': 'serif',
    'font.serif': ['Georgia', 'DejaVu Serif', 'Times New Roman'],
    'font.size': 14,
    'axes.labelsize': 15,
    'axes.titlesize': 18,
    'xtick.labelsize': 12,
    'ytick.labelsize': 12,
})


def chart1_prevalence():
    """Slide 2: AD Prevalence Projection – line chart 7.2M (2025) to 13.8M (2060)."""
    years = [2025, 2030, 2035, 2040, 2045, 2050, 2055, 2060]
    prevalence = [7.2, 8.0, 9.1, 10.3, 11.2, 12.1, 13.0, 13.8]

    fig, ax = plt.subplots(figsize=(8, 4.5))
    fig.patch.set_alpha(0)
    ax.set_facecolor('none')

    ax.plot(years, prevalence, color=COPPER, linewidth=3, marker='o',
            markersize=8, markerfacecolor=COPPER, markeredgecolor='white', markeredgewidth=1.5)
    ax.fill_between(years, prevalence, alpha=0.12, color=COPPER)

    # Annotate endpoints
    ax.annotate('7.2M', xy=(2025, 7.2), xytext=(2025, 7.8),
                fontsize=16, fontweight='bold', color=COPPER, ha='center')
    ax.annotate('13.8M', xy=(2060, 13.8), xytext=(2060, 14.4),
                fontsize=16, fontweight='bold', color=RED, ha='center')

    ax.set_xlabel('Year', color=TEXT_DARK, fontweight='semibold')
    ax.set_ylabel('Americans with AD (millions)', color=TEXT_DARK, fontweight='semibold')
    ax.set_title('Projected Alzheimer\'s Prevalence', color=TEXT_DARK, fontweight='bold', pad=15)

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(CITE_GRAY)
    ax.spines['bottom'].set_color(CITE_GRAY)
    ax.tick_params(colors=TEXT_DARK)
    ax.set_ylim(5, 16)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.1f'))

    plt.tight_layout()
    fig.savefig(os.path.join(OUTPUT_DIR, "prevalence_projection.png"),
                dpi=300, transparent=True, bbox_inches='tight')
    plt.close(fig)
    print("  ✓ Chart 1: prevalence_projection.png")


def chart2_ml_comparison():
    """Slide 16: Classical ML Comparison – grouped bars AD-vs-HC (~94%) vs MCI (~70%)."""
    categories = ['AD vs HC', 'MCI vs HC']
    svm   = [94.5, 68.0]
    rf    = [91.0, 65.0]
    lr    = [89.5, 63.0]

    x = np.arange(len(categories))
    width = 0.22

    fig, ax = plt.subplots(figsize=(8, 4.5))
    fig.patch.set_alpha(0)
    ax.set_facecolor('none')

    bars1 = ax.bar(x - width, svm, width, label='SVM', color=COPPER, edgecolor='white', linewidth=0.5)
    bars2 = ax.bar(x, rf, width, label='Random Forest', color=BLUE, edgecolor='white', linewidth=0.5)
    bars3 = ax.bar(x + width, lr, width, label='Logistic Regression', color=GREEN, edgecolor='white', linewidth=0.5)

    # Value labels
    for bars in [bars1, bars2, bars3]:
        for bar in bars:
            h = bar.get_height()
            ax.annotate(f'{h:.1f}%', xy=(bar.get_x() + bar.get_width()/2, h),
                        xytext=(0, 4), textcoords='offset points',
                        ha='center', fontsize=10, fontweight='bold', color=TEXT_DARK)

    ax.set_ylabel('Accuracy (%)', color=TEXT_DARK, fontweight='semibold')
    ax.set_title('Classical ML: AD Detection vs MCI Detection', color=TEXT_DARK, fontweight='bold', pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontweight='semibold', color=TEXT_DARK)
    ax.set_ylim(0, 110)
    ax.legend(frameon=False, fontsize=11)

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(CITE_GRAY)
    ax.spines['bottom'].set_color(CITE_GRAY)
    ax.tick_params(colors=TEXT_DARK)

    # Draw MCI "cliff" annotation
    ax.annotate('', xy=(1.3, 70), xytext=(1.3, 94),
                arrowprops=dict(arrowstyle='<->', color=RED, lw=2))
    ax.text(1.45, 82, '~25% drop', fontsize=11, color=RED, fontweight='bold', va='center')

    plt.tight_layout()
    fig.savefig(os.path.join(OUTPUT_DIR, "ml_comparison.png"),
                dpi=300, transparent=True, bbox_inches='tight')
    plt.close(fig)
    print("  ✓ Chart 2: ml_comparison.png")


def chart3_roc_curve():
    """Slide 19: Conceptual ROC Curve – good classifier vs chance."""
    fpr_good = np.array([0, 0.02, 0.05, 0.10, 0.15, 0.25, 0.40, 0.60, 1.0])
    tpr_good = np.array([0, 0.45, 0.65, 0.78, 0.85, 0.92, 0.96, 0.98, 1.0])

    fig, ax = plt.subplots(figsize=(5.5, 5.5))
    fig.patch.set_alpha(0)
    ax.set_facecolor('none')

    # Chance line
    ax.plot([0, 1], [0, 1], linestyle='--', color=CITE_GRAY, linewidth=1.5, label='Chance (AUC = 0.50)')

    # Good classifier
    ax.fill_between(fpr_good, tpr_good, alpha=0.10, color=BLUE)
    ax.plot(fpr_good, tpr_good, color=BLUE, linewidth=3, label='Best Model (AUC = 0.93)')

    # Optimal point
    ax.scatter([0.10], [0.78], color=COPPER, s=120, zorder=5, edgecolors='white', linewidths=2)
    ax.annotate('Optimal\nThreshold', xy=(0.10, 0.78), xytext=(0.25, 0.65),
                fontsize=12, fontweight='bold', color=COPPER,
                arrowprops=dict(arrowstyle='->', color=COPPER, lw=1.5))

    ax.set_xlabel('False Positive Rate', color=TEXT_DARK, fontweight='semibold')
    ax.set_ylabel('True Positive Rate', color=TEXT_DARK, fontweight='semibold')
    ax.set_title('ROC Curve: AD Classification', color=TEXT_DARK, fontweight='bold', pad=15)
    ax.legend(loc='lower right', frameon=False, fontsize=11)

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(CITE_GRAY)
    ax.spines['bottom'].set_color(CITE_GRAY)
    ax.tick_params(colors=TEXT_DARK)
    ax.set_xlim(-0.02, 1.02)
    ax.set_ylim(-0.02, 1.05)

    plt.tight_layout()
    fig.savefig(os.path.join(OUTPUT_DIR, "roc_curve.png"),
                dpi=300, transparent=True, bbox_inches='tight')
    plt.close(fig)
    print("  ✓ Chart 3: roc_curve.png")


def chart4_data_leakage():
    """Slide 23: Data Leakage Impact – two bars: with leakage (~95%) vs without (~67%)."""
    labels = ['With Data\nLeakage', 'Without Data\nLeakage']
    values = [95, 67]
    colors = [RED, GREEN]

    fig, ax = plt.subplots(figsize=(6, 4.5))
    fig.patch.set_alpha(0)
    ax.set_facecolor('none')

    bars = ax.bar(labels, values, width=0.5, color=colors, edgecolor='white', linewidth=1)

    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1.5,
                f'{val}%', ha='center', fontsize=20, fontweight='bold',
                color=bar.get_facecolor())

    # Arrow showing the drop
    ax.annotate('', xy=(1, 67), xytext=(0, 95),
                arrowprops=dict(arrowstyle='->', color=TEXT_DARK, lw=2.5,
                                connectionstyle='arc3,rad=-0.2'))
    ax.text(0.5, 85, '−28%', fontsize=22, fontweight='bold', color=RED,
            ha='center', va='center',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor=RED, alpha=0.9))

    ax.set_ylabel('Reported Accuracy (%)', color=TEXT_DARK, fontweight='semibold')
    ax.set_title('Impact of Data Leakage on Model Performance',
                 color=TEXT_DARK, fontweight='bold', pad=15)
    ax.set_ylim(0, 110)

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(CITE_GRAY)
    ax.spines['bottom'].set_color(CITE_GRAY)
    ax.tick_params(colors=TEXT_DARK)

    plt.tight_layout()
    fig.savefig(os.path.join(OUTPUT_DIR, "data_leakage_impact.png"),
                dpi=300, transparent=True, bbox_inches='tight')
    plt.close(fig)
    print("  ✓ Chart 4: data_leakage_impact.png")


if __name__ == "__main__":
    print("Generating presentation charts...")
    chart1_prevalence()
    chart2_ml_comparison()
    chart3_roc_curve()
    chart4_data_leakage()
    print(f"\nAll charts saved to: {OUTPUT_DIR}")
