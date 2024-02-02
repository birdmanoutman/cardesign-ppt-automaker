import openai
openai.api_key= r'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'
completion = openai.ChatCompletion.create(
    model='gpt-3.5-turbo',
    messages=[{"role": "user", 'content':"how are you"}]
)

print(completion)