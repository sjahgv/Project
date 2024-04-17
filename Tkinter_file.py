import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import SparkApi
import json

# 以下密钥信息从控制台获取
appid = "935ff112"  # 填写控制台中获取的 APPID 信息
api_secret = "MDRlMDVkNzc2MDBjNTUzMzQ4ODFjZmM3"  # 填写控制台中获取的 APISecret 信息
api_key = "2631b14903cc6f29c0b914e19bc46987"  # 填写控制台中获取的 APIKey 信息

# 用于配置大模型版本，默认“general/generalv2”
domain = "generalv3"  # v1.5版本
# domain = "generalv2"    # v2.0版本
# 云端环境的服务地址
Spark_url = "ws://spark-api.xf-yun.com/v3.1/chat"  # v1.5环境的地址


def read_excel(excel_path, start_row, end_row, column_name):
    # 读取Excel文件
    df = pd.read_excel(excel_path)

    # 确保指定的列名存在于DataFrame中
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' does not exist in the DataFrame.")

    # 使用iloc来指定行号范围并读取特定列
    # 从第start_row行开始，到第end_row行结束
    specific_column_content = df.iloc[start_row - 2:end_row - 1][column_name].values

    return specific_column_content


text = []


def getText(role, content):
    jsoncon = {}
    jsoncon["role"] = role
    jsoncon["content"] = content
    text.append(jsoncon)
    return text

def process_excel():
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if excel_path:
        start_row = start_row_entry.get()
        end_row = end_row_entry.get()
        column_name = column_name_entry.get()

        # 检查输入框是否为空
        if not start_row or not end_row or not column_name:
            messagebox.showerror("错误", "请确保所有输入框都有值。")
            return  # 退出函数

        try:
            start_row = int(start_row)
            end_row = int(end_row)
        except ValueError:
            messagebox.showerror("错误", "请输入有效的整数值。")
            return  # 退出函数
        content = read_excel(excel_path, start_row, end_row, column_name)
        result = ''
        # 设置一个提示标签，在处理Excel文件时显示“请稍等，正在识别”
        processing_label = tk.Label(app, text="请稍等，正在识别")
        processing_label.pack(pady=10)
        for item in content:
            result += str(item)
        que = ('''请依据《国际中文教育中文水平等级标准》中的“四维基准”，并结合老师所强调的内容，从以下教学对话中自动识别并提取重点词语、语音、语法或话题知识点，尤其是其运用。
                让我们一步一步思考，按照以下内容进行：
                1. 理解任务要求：首先，你需要从一段教学对话中识别和提取词语、语法、语音和话题知识点。
                2. 分析对话内容：接下来，分析提供的教学对话内容，寻找关键词语、语法结构、语音特点和话题线索。在这个过程中，要注意到老师和学生之间的互动，以及老师如何引导学生理解和使用特定的词汇和语法结构。
                3. 提取关键词汇：从对话中提取关键词汇或被特别强调的词汇。
                4. 识别语法点：注意到汉语中的重要语法点。
                5. 考虑语音特征：留意发音、声调等语音方面的讨论。
                6. 确定话题：话题主要是围绕学习词汇和语法，这些内容在对话中被反复提及。
                7. 结果组织: 将知识点按照JSON呈现。
                输出形式如下：
                {
                  "词汇": {
                    "葡萄酒"
                      },
                  "语法": {
                    "把字句": ["把一瓶酒喝完", "把书打开"],
                   },
                  "语音": {
                    "可乐": ["发音不准"]
                  },
                  "话题": {
                       "做客": ["你去朋友家做客的时候，你常常带什么礼物？"],  }
                }

                以下是需要你识别的内容：''') + result

        question = getText("user", que)

        # question = "以下内容中老师强调了哪些知识点？" + content
        print(question)
        SparkApi.answer = ""
        print("星火:", end="")

        SparkApi.main(appid, api_key, api_secret, Spark_url, domain, question)
        if SparkApi.answer == "":
            messagebox.showerror("出错", "未找到内容")
        else:

            # result_text.delete('1.0', tk.END)  # 清空之前的文本
            # result_text.insert(tk.END, f"识别结果如下\n{SparkApi.answer}\n")  # 插入新的结果文本
            result_text.insert(tk.END, f"识别结果如下\n{SparkApi.answer}\n")  # 插入新的结果文本
            messagebox.showinfo("完成", "知识点识别完成，结果已显示。")
            # 添加保存结果的逻辑
            save_button = tk.Button(frame, text="保存结果", command=lambda: save_results())
            save_button.grid(row=4, columnspan=2, pady=10)

        # 处理完成后清除提示信息
        # SparkApi.answer = ""
        processing_label.config(text="")
def save_results():
    # 将文本框中的内容保存到文件
    with open('output.txt', 'w', encoding='utf-8') as file:
        file.write(result_text.get('1.0', tk.END))
    messagebox.showinfo("保存完成", "结果已保存到输出文件中。")



# GUI界面设计
# 创建一个Tkinter窗口对象
app = tk.Tk()
app.title("课堂教学内容知识点识别")
# 创建一个框架容器，用于布局其他控件
frame = tk.Frame(app)
# 设置框架容器在窗口中的位置和填充
frame.pack(pady=20)
# 创建标签和输入框，用于用户输入起始行号、结束行号和列名
start_row_label = tk.Label(frame, text="起始行号：")
start_row_label.grid(row=0, column=0, sticky="e")  # 将标签放置在框架中，并设置对齐方式
start_row_entry = tk.Entry(frame)  # 创建一个输入框
start_row_entry.grid(row=0, column=1)  # 将输入框放置在框架中

end_row_label = tk.Label(frame, text="结束行号：")
end_row_label.grid(row=1, column=0, sticky="e")
end_row_entry = tk.Entry(frame)
end_row_entry.grid(row=1, column=1)

column_name_label = tk.Label(frame, text="列名：")
column_name_label.grid(row=2, column=0, sticky="e")
column_name_entry = tk.Entry(frame)
column_name_entry.grid(row=2, column=1)
# 创建一个文本框，用于显示处理后的结果
result_text = tk.Text(app, height=30, width=100)  # 创建一个文本框对象
result_text.pack(pady=10)  # 将文本框放置在窗口中，并设置填充

process_button = tk.Button(frame, text="处理Excel", command=process_excel)
process_button.grid(row=3, columnspan=2, pady=10)


app.mainloop()
