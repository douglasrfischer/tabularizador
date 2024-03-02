class Config(object):
    '''Configuração do serviço em produção.'''
    SECRET_KEY = 'KEY_PLACEHOLDER'
    UPLOAD_FOLDER = 'Diretório\\Da\\Aplicação\\tabularizador\\application\\uploads'
    DOWNLOAD_FOLDER = 'C:\\ApplicationDirectory\\tabularizador\\application\\downloads'
    ALLOWED_FORMATS = {'xlsx'}
