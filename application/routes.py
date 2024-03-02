from flask import render_template, request, redirect, flash, Blueprint, send_file
from flask import current_app as app
from werkzeug.utils import secure_filename
from .tabularizacao import tabularizar
from os import path


public_bp = Blueprint('public_bp', __name__)

@public_bp.app_errorhandler(404)
def not_found(e):
    return '<h2>404: página não encontrada.</h2>'


@public_bp.app_errorhandler(500)
def internal_server_error(e):
    return '<h2>500: erro interno do servidor. Contate o administrador</h2>'


@public_bp.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

# ------------------------------------------------------

def arquivo_permitido(arquivo) -> bool:
    return arquivo.filename.split('.')[-1] in app.config['ALLOWED_FORMATS']


@public_bp.post('/upload')
def upload():
    arquivo = request.files['arquivo']

    if not arquivo:
        flash('Nenhum arquivo carregado.')
        return redirect('/')
    
    if arquivo_permitido(arquivo):
        arquivo.save(
            path.join(app.config['UPLOAD_FOLDER'], secure_filename(arquivo.filename)
            ))
    else:
        flash('Formato não suportado.')
        return redirect('/')

    try:
        arquivo_nome_tratado = arquivo.filename.replace(' ', '_')
        nome = tabularizar(arquivo_nome_tratado)
    except Exception as e:
        flash('Algo deu errado! Verifique a planilha e o nome da aba (deve ser "LL ABERTO (T2T)")')
        print(f'\n\n{e}\n\n')
        return redirect('/')
    
    return redirect(f'/download/{nome}')


@public_bp.get('/download/<nome>')
def download(nome):
    return send_file(
        path_or_file=f'{app.config["DOWNLOAD_FOLDER"]}\\{nome}',
        as_attachment=True
        )
