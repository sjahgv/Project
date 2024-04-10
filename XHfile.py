import pandas as pd
import requests
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
    specific_column_content = df.iloc[start_row -2:end_row -1][column_name].values

    return specific_column_content

text=[]
def getText(role,content):
    jsoncon = {}
    jsoncon["role"] = role
    jsoncon["content"] = content
    text.append(jsoncon)
    return text
# 主函数
def main():
    #excel_path = './标注20-1.xlsx'
    # 读取的是416-445，excel行号
    # start_row = 353  # 起始行号（从1开始）
    # end_row = 434  # 结束行号（从1开始）
    # column_name = '话语'  # 替换为你的列名
    i = 0

    excel_path = './' + input("请输入Excel文件名（包括扩展名，例如：data.xlsx）: ")
    with open('output.txt', 'w', encoding='utf-8') as file:
        # 初始化循环标志
        continue_flag = 'y'
        while continue_flag.lower() == 'y':
            i = i+1

            start_row =int(input("请输入起始行号: "))
            end_row = int(input("请输入结束行号: "))
            column_name = input("请输入需要读取列的列名: ")
            # 调用函数
            content = read_excel(excel_path, start_row, end_row, column_name)
            result = ''
            for item in content:
                result += str(item)
            que= ('''请依据《国际中文教育中文水平等级标准》中的“四维基准”，并结合老师所强调的内容，从以下教学对话中自动识别并提取重点词语、语音、语法或话题知识点，尤其是其运用。
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

            question =getText("user",que)

            # question = "以下内容中老师强调了哪些知识点？" + content
            print(question)
            SparkApi.answer = ""
            print("星火:", end="")
            SparkApi.main(appid, api_key, api_secret, Spark_url, domain, question)
        # print(SparkApi.answer)
        # 将数组转换为JSON格式的字符串# 设置sort_keys=True来确保键的顺序一致
        # 打开一个文件用于写入，如果文件不存在将会被创建
        #     with open('output.txt', 'w', encoding='utf-8') as file:

            # 将JSON字符串写入文件
            file.write("第" + str(i) + "次识别" + "\n"+SparkApi.answer + "\n")
            # 询问用户是否还有内容需要识别
            continue_flag = input("是否还有内容需要识别？(y/n): ")
    print("内容识别完成")

    # print(content)





if __name__ == '__main__':
    main()

