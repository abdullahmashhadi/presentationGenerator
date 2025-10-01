from app import app

# Vercel serverless function handler
def handler(request, response):
    return app(request, response)

if __name__ == '__main__':
    app.run(debug=False)
