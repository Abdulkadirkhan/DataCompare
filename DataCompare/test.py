import requests as req

data =  req.get('https://jsonplaceholder.typicode.com/todos/1')

print('chumma chumma ', data.text)
