import json
import requests
import io
import base64
from PIL import Image, PngImagePlugin

url = "http://124.221.249.173:41430"
#
# payload = {
#     "prompt": "puppy dog, cute, nice",
#     "steps": 5
# }
#
# response = requests.post(url=f'{url}/sdapi/v1/txt2img', json=payload)
#
# r = response.json()
# print(r)
#
# for i in r['images']:
#     image = Image.open(io.BytesIO(base64.b64decode(i.split(",",1)[0])))
#     image.save('output.png')


imgpath = r'/Users/birdmanoutman/Desktop/pokemon.jpeg'
with open(imgpath, 'rb') as imgfile:
    imgdata = imgfile.read()

encoded_img = base64.b64encode(imgdata).decode('utf-8')

print(encoded_img)

payload2 = {
    "image": encoded_img,
    "model": "deepbooru"
}
response2 = requests.post(url=url+'/sdapi/v1/interrogate', json=payload2)
print(response2.json())

