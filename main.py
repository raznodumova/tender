from flask import Flask, request, send_file, render_template
import pandas as pd
import requests
import os

app = Flask(__name__)

VK_API_URL = "https://vk.cc/shorten"
ACCESS_TOKEN = " "  # Вставьте ваш токен VK


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file uploaded!", 400

    file = request.files['file']

    if file.filename.endswith('.xlsx'):
        df = pd.read_excel(file)
        if df.shape[1] < 1:
            return "No links found in the file!", 400

        # Предполагаем, что ссылки находятся в первом столбце
        links = df.iloc[:, 0].tolist()
        short_links = [shorten_link(link) for link in links]

        df['Short Link'] = short_links
        output_file = f"shortened_links.xlsx"
        df.to_excel(output_file, index=False)

        return send_file(output_file, as_attachment=True)
    else:
        return "Invalid file format! Please upload an .xlsx file.", 400


def shorten_link(link):
    params = {'url': link}
    try:
        response = requests.post(VK_API_URL, params=params, headers={'Authorization': f'Bearer {ACCESS_TOKEN}'})
        if response.status_code == 200:
            return response.json().get('short_url', link)
        else:
            return link  # Вернуть оригинальную ссылку в случае ошибки
    except Exception as e:
        return link  # Вернуть оригинальную ссылку в случае исключения


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
