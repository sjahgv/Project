# import zhipuai
#
# zhipuai.api_key = "e93586d257e90b7ec52b0a277a9a3128.aQvxs94yGEDX8exg"  # 填写控制台中获取的 APIKey 信息
# model = "glm-4"  # 用于配置大模型版本
#
#
# def getText(role, content, text=[]):
#     # role 是指定角色，content 是 prompt 内容
#     jsoncon = {}
#     jsoncon["role"] = role
#     jsoncon["content"] = content
#     text.append(jsoncon)
#     return text
#
# question = getText("user", "你好")
# print(question)
# # 请求模型
# response = zhipuai.model_api.invoke(
#     model=model,
#     prompt=question
# )
# print(response)


from zhipuai import ZhipuAI
client = ZhipuAI(api_key="e93586d257e90b7ec52b0a277a9a3128.aQvxs94yGEDX8exg") # 填写您自己的APIKey
# response = client.chat.completions.create(
#     model="glm-4",  # 填写需要调用的模型名称
#     messages=[
#         {"role": "user", "content": "作为一名营销专家，请为我的产品创作一个吸引人的slogan"},
#     ],
# )
# print(response.choices[0].message)
prompt = ""
while True:
    prompt = input("user:")
    response = client.chat.completions.create(
        model="glm-4",  # 填写需要调用的模型名称
        messages=[
            {"role": "user", "content": "你好"},
            {"role": "assistant", "content": "我是人工智能助手"},
            {"role": "user", "content": "你叫什么名字"},
            {"role": "assistant", "content": "我叫chatGLM"},
            {"role": "user", "content": prompt}
        ],
    )
    answer = response.choices[0].message
    print(dict(answer)["content"])