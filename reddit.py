import requests
import json
import xlsxwriter

class Post:
    def __init__(self,title,body,url,flair,ups,downs):
        self.title = title
        self.body = body
        self.flair = flair
        self.votes = Votes(ups,downs)
        self.url = url

    def printPost(self):
        print(self.url)

class Votes:
    def __init__(self, up, down):
        self.up = up
        self.down = down
    def toJSON(self):
        return json.dumps(self, default=lambda o: o.__dict__, 
            sort_keys=True, indent=4)



class Comment:
    def __init__(self,body,url,ups,downs):
        self.body = body
        self.url = url
        self.votes = Votes(ups,downs)
    def toJSON(self):
        return json.dumps(self, default=lambda o: o.__dict__, 
            sort_keys=True, indent=4)

def getPosts():
    url = "https://www.reddit.com/r/Indian_Academia/top.json?limit=20"
    payload={}
    headers = {
    'User-agent': 'mybot 0.0',
    }
    response = requests.request("GET", url, headers=headers, data=payload)

    responseData = response.json()

    data = responseData['data']

    posts = []

    for child in data["children"]:
        childData = child["data"]
        title = childData["title"]
        ups = childData['ups']
        downs = childData['downs']
        flair = childData['link_flair_text']
        url = childData['url']
        body = childData['selftext']
        post = Post(title,body,url,flair,ups,downs)
        posts.append(post)
    
    return posts


def getComments(postUrl):
    postJson = postUrl +".json"
    payload={}
    headers = {
    'User-agent': 'mybot 0.1',
    }

    for i in range(3):
        response = requests.request("GET", postJson, headers=headers, data=payload)
        if 200 == response.status_code:
            responseData = response.json()
            break
        else:
            print("status code while getComments() is %d" %(response.status_code))

    

    ## check index out of bound
    replyData = responseData[1]
    replies = recursivelyGetReplies(replyData['data'])
    return replies




def recursivelyGetReplies(data):
    replyList = []
    for reply in data['children']:
        commentData = reply['data']
        if 'ups' not in commentData:
            return replyList
        ups = commentData['ups']
        downs = commentData['downs']
        body = commentData['body']
        url = "https://www.reddit.com" + commentData['permalink']
        replyList.append(Comment(body,url,ups,downs))
        replies = commentData['replies']
        if replies != "":
            rList = recursivelyGetReplies(replies['data'])
            replyList.extend(rList)
    return replyList

outWorkbook = xlsxwriter.Workbook("out.xlsx")
outSheet = outWorkbook.add_worksheet()

outSheet.write(0,0,"Post URL")
outSheet.write(0,1,"Post Title")
outSheet.write(0,2,"Comment URL")
outSheet.write(0,3,"Comment Content")

row = 1
col = 0

posts = getPosts()

for post in posts:
    replies = getComments(post.url)
    for reply in replies:
        outSheet.write(row,col,post.url)
        outSheet.write(row,col+1,post.title)
        outSheet.write(row,col+2,reply.url)
        outSheet.write(row,col+3,reply.body)
        row = row + 1



outWorkbook.close()