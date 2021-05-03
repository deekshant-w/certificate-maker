import base64
import json

# settings = {
#     "emailFrom"     :"deekshantwadhwa2000@gmail.com",
#     "inputData"      :"ldqsjwdxgkzqimsv"
# }

settings = {
    "emailFrom"     :"a@gmail.com",
    "inputData"     :"a"
}


storeVal = json.dumps(settings).encode('ascii')

result = base64.b64encode(base64.b64encode(base64.b64encode(base64.b64encode(base64.b64encode(storeVal)))))
handle = open("input.dk",'wb')
handle.write(result)
handle.close()

handle2 = open("input.dk",'rb').read()

dec = (base64.b64decode(base64.b64decode(base64.b64decode(base64.b64decode(base64.b64decode(handle2)))))).decode('utf8').replace("'", '"')
data = json.loads(dec)
