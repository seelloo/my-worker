from openai import OpenAI
import os

# 从环境变量中获取 API-KEY
# 这是一个更安全的做法，避免将密钥硬编码到代码中
# 建议你在运行代码前设置环境变量：
# export QWEN_API_KEY="你的API-KEY"
api_key = os.getenv("QWEN_API_KEY")

if not api_key:
    raise ValueError("请设置 QWEN_API_KEY 环境变量")

# 创建 OpenAI 客户端实例
# baseURL 指向阿里云的 Qwen 服务地址
client = OpenAI(
    api_key='sk-',
    base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
)

def get_qwen_response(prompt):
    """向通义千问发送请求并获取响应"""
    try:
        completion = client.chat.completions.create(
            model="qwen-turbo",  # 或 "qwen-plus", "qwen-max" 等模型
            messages=[
                {
   "role": "user", "content": prompt}
            ],
            temperature=0.8,  # 控制回答的创造性，0.0-1.0
        )
        return completion.choices[0].message.content
    except Exception as e:
        print(f"调用 API 失败: {e}")
        return None

# 调用函数并打印结果
user_prompt = "帮我写一个Python函数，用于计算斐波那契数列。"
response = get_qwen_response(user_prompt)

if response:
    print("--- Qwen 的回答 ---")
    print(response)
