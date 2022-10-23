from matplotlib.font_manager import FontProperties
import matplotlib.pyplot as plt
import numpy as np


# 繪製長條圖
def plot_bar(data={'apple': 10, 'orange': 15, 'lemon': 5, 'lime': 20},
             width=0.6, fontsize=6, fname='plotname', digit=True):
    unsort_names, unsort_values = list(data.keys()), list(map(float, data.values()))

    unsort_dict = dict(zip(unsort_names, unsort_values))
    sort_dict = dict(sorted(unsort_dict.items(), key=lambda item: item[1], reverse=True))

    names, values = list(sort_dict.keys()), list(sort_dict.values())
    x = np.arange(len(names))  # the label locations

    fig, ax = plt.subplots()
    bar = ax.bar(x, values, width=width, color='#19b8be')  # 設定長條圖圖型

    # 設定標籤
    # ax.set_xticks(x, names, rotation=90)  # 設定X軸文字
    # ax.set_yticklabels(['{:,.0f}'.format(x) for x in ax.get_yticks()])  # 設定Y軸文字

    ax.set_xticks(x)  # values
    ax.set_xticklabels(names, rotation=90)  # labels


    # 取消右方、上方與左方邊界
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['left'].set_visible(False)
    # ax.yaxis.set_ticks_position('left')
    ax.xaxis.set_ticks_position('bottom')
    ax.get_yaxis().set_visible(False)  # 隱藏y軸

    # 設定長條圖上方數字標註
    label_diff, label_growth = '{:,.0f}', '{:,.1f}'
    label_style = label_growth if digit == True else label_diff
    ax.bar_label(bar, labels=[label_style.format(x) for x in values], padding=3,
                 fontsize=fontsize, color='#19b8be')

    # 使用 for 迴圈一一取出 x 軸標籤 label 設定字體，若 y 軸有中文字也是類似使用方式 get_yticklabels
    myfont = FontProperties(fname=r'./plot/TaipeiSansTCBeta-Regular.ttf', size=12)
    for label in ax.get_xticklabels():
        label.set_fontproperties(myfont)

    # 增加ylim的lower bound
    xmin, xmax, ymin, ymax = plt.axis()
    add_ymin = ymin - np.add(abs(ymax), abs(ymin))/10
    plt.ylim([add_ymin, ymax])

    fig.tight_layout()
    plt.savefig("plot/" + fname, bbox_inches="tight", dpi=300)
    print('save ', fname, ' as a jpg')
    # plt.show()


if __name__ == '__main__':
    pass
