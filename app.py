from flask import Flask, send_from_directory

app = Flask(__name__, static_folder='static')

# 根路径返回静态 index.html
@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

# 其它静态资源（.js/.css/.json/.png …）
@app.route('/<path:path>')
def static_proxy(path):
    return send_from_directory('static', path)

# Heroku 会读取 PORT 环境变量
if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
