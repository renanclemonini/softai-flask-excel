from flask import Flask, request, render_template, redirect, url_for, send_from_directory, flash
from utils import allowed_file, converter_xls_para_xlsx, processar_excel_oficial3, limpar_pasta_input
from werkzeug.utils import secure_filename
from uuid import uuid4
import os
import time
import secrets


api_key = 'bafbb1fc22b54b6a92e08f877b2a80d507dc65a2f8b0936eebc986c4c3a03aa2'


def create_app():
    app = Flask(__name__)
    app.secret_key = api_key
    limpar_pasta_input() 

    @app.route('/')
    def index():
        return redirect(url_for('upload'))

    @app.route('/upload', methods=['GET', 'POST'])
    def upload():
        if request.method == 'POST':
            file = request.files['file']
            if file and allowed_file(file.filename):
                try:
                    filename = secure_filename(file.filename)
                    os.makedirs('input', exist_ok=True)
                    save_location = os.path.join('input', filename)
                    file.save(save_location)

                    if save_location.endswith('.xls'):
                        save_location = converter_xls_para_xlsx(save_location)

                    # arquivo_processado = processar_excel_oficial2(save_location)
                    response = processar_excel_oficial3(save_location)
                    try:
                        os.remove(save_location)
                    except Exception as e:
                        print(f"Erro ao deletar o arquivo: {e}")

                    # arquivo_nome = os.path.basename(response.arquivo_processado)
                    arquivo_nome = os.path.basename(response['arquivo_processado'])
                    return render_template('download.html', arquivo_dowload=arquivo_nome, response_data=response)
                except Exception as e:
                    flash("Arquivo corrompido!", "warning")
                    return redirect(url_for('upload'))

        return render_template('input_page.html')
    
    @app.route('/dowload')
    def download():
        return render_template('dowload.html')
    
    @app.route('/download/<filename>')
    def download_file(filename):
        return send_from_directory('output', filename, as_attachment=True)

    return app


if __name__ == '__main__':
    app = create_app()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
