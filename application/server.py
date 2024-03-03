import flask

server = flask.Flask(__name__)

@server.get("/")
def index():
    return flask.render_template("index.jinja")

@server.get("/taskpane")
def taskbar():
    return flask.render_template("taskpane.jinja")


server.debug = True
server.run(port=3000, host="0.0.0.0")
