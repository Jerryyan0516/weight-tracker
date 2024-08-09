import pandas as pd
from datetime import datetime
import os
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.dates as mdates
import tkinter as tk
from tkinter import simpledialog, messagebox

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

file_path = r'C:\Users\65179\Desktop\每日体重记录\date_weight.xlsx'

def record_weight(weight, file_path):
    """
    记录体重数据并写入到 Excel 文件中，同时与上次体重比较，给出提示。
    """
    # 获取当前日期和时间
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # 创建一个包含当前时间和体重的新数据
    new_data = pd.DataFrame({'时间': [current_time], '体重(kg)': [weight]})

    # 检查文件是否存在
    if os.path.exists(file_path):
        # 如果文件存在，读取现有数据
        df = pd.read_excel(file_path)

        # 将时间列转换为datetime格式，避免数据类型混淆
        df['时间'] = pd.to_datetime(df['时间'], errors='coerce')
        new_data['时间'] = pd.to_datetime(new_data['时间'], errors='coerce')

        # 获取最后一条记录的体重
        last_weight = df.iloc[-1]['体重(kg)']

        # 比较体重变化
        weight_diff = weight - last_weight
        if weight_diff > 0:
            message = f"比上次增加了 {weight_diff:.2f} kg，建议增加运动量！"
        elif weight_diff < 0:
            message = f"比上次减少了 {abs(weight_diff):.2f} kg，继续加油！"
        else:
            message = "体重没有变化，保持良好的习惯！"
        
        # 显示结果的弹窗
        messagebox.showinfo("体重变化结果", message)

        # 将新数据追加到现有数据中
        df = pd.concat([df, new_data], ignore_index=True)
    else:
        # 如果文件不存在，创建一个新的DataFrame
        df = new_data

    # 按时间列进行排序
    df = df.sort_values(by='时间')

    # 将排序后的数据写回 Excel 文件
    df.to_excel(file_path, index=False)



def plot_weight_trend(file_path):
    """
    读取 Excel 文件并生成体重变化的折线图。
    """
    # 读取 Excel 文件
    df = pd.read_excel(file_path)

    # 将时间列转换为datetime格式
    df['时间'] = pd.to_datetime(df['时间'])

    # 绘制折线图
    plt.figure(figsize=(20, 12))
    plt.plot(df['时间'], df['体重(kg)'], marker='o', linestyle='-', color='b')

    # 设置图表标题和轴标签
    plt.title('yy的体重变化折线图', fontsize=18)
    plt.xlabel('日期', fontsize=22)
    plt.ylabel('体重(kg)', fontsize=22)

    # 设置日期格式
    date_format = mdates.DateFormatter('%Y-%m-%d %H:%M:%S')
    plt.gca().xaxis.set_major_formatter(date_format)

    # 调整日期标签的显示方式
    plt.gcf().autofmt_xdate()


    # 旋转横坐标标签以更好地显示日期
    plt.xticks(rotation=45)

    # 添加网格
    plt.grid(True)

    # 显示图表
    plt.tight_layout()  # 自动调整子图参数，使之填充整个图像区域
    plt.show()


def main():
    """
    主函数，用于执行体重记录和折线图绘制。
    """
    # 创建一个Tkinter窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 弹出一个输入框让用户输入体重
    weight_input = simpledialog.askstring("体重记录", "请输入今天的体重(kg) (或按 Enter 直接查看图表):")

    # 如果输入为空字符串，则直接展示图表
    if weight_input is None or weight_input.strip() == '':
        plot_weight_trend(file_path)
    else:
        weight = float(weight_input)
        record_weight(weight, file_path)
        plot_weight_trend(file_path)


# 调用主函数
if __name__ == "__main__":
    main()
