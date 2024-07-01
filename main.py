from flask import Flask, render_template, url_for, request, flash, send_from_directory
from forms import FormSalvar,FormInfoVazamentos
from werkzeug.utils import secure_filename
import os,shutil
import f_qualidade.tratar_dados_qualidade,f_qualidade.planilha_qualidade
import f_eficiencia.tratar_dados_eficiencia,f_eficiencia.planilha_eficiencia
import f_vazamentos.tratar_dados_vazamentos
import webbrowser, time
import keyboard

chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe %s --incognito'
webbrowser.get(chrome_path).open_new('http://127.0.0.1:5000')

ALLOWED_EXTENSIONS = {'xlsx','zip'}
UPLOAD_FOLDER = 'arquivos'

dados_dict = {'Tabela de dados':['Vazio']}
results_dict = {}
dados_empresa = {
    'Empresa': '',
    'CNPJ': '',
    'Endereço': '',
    'Contato': '',
    'Departamento': '',
    'E-mail': '',
    'Telefone': '',
    'RT': '',
    'Membro 1': '',
    'Membro 2': '',
    'Membro 3': '',
    'Membro 4': '',
}
flag_registrado = False
flag_ready = False
flag_atividade = ''
nome_arquivo = ''

app = Flask(__name__)

app.config['SECRET_KEY'] = '358823e5046ab23c149ff9a047b30ae8'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

#Exclusão de arquivos após concluir ação
def limpar_pasta(folder):
    for file in os.listdir(folder):
        if 'modelo' not in file and '.txt' not in file:
            try:
                os.remove(f'{folder}/{file}')
            except:
                shutil.rmtree(f'{folder}/{file}')
#--------------------------------------------------------------------------------------------------------


#Página de qualidade de energia
@app.route("/qualidade",methods=['GET','POST'])
def qualidade():
    global dados_dict
    global flag_ready
    global results_dict
    global nome_arquivo
    global flag_atividade
    form_salvar = FormSalvar()

    limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))

    #Procedimento para carregar arquivo
    if request.method == 'POST' and 'load_btn' in request.form:     
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            dados_dict = f_qualidade.tratar_dados_qualidade.carregar_dados(file = filename,folder=UPLOAD_FOLDER)
            results_dict = f_qualidade.tratar_dados_qualidade.calculos(dados_dict=dados_dict)
            flag_ready = True
            flash('Dados carregados com sucesso',category='alert-success')
            return app.redirect(url_for('qualidade'))
    #--------------------------------------------------------------------------------------------------------
    
    #Procedimento para salvar a planilha com as análises
    if 'salvar_btn' in request.form:            
        msg = f_qualidade.tratar_dados_qualidade.verificar_save(dados_dict=dados_dict)
        if msg != 'Arquivo salvo com sucesso': 
            flash(msg,category='alert-danger')
        else:
            limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))
            nome_arquivo = form_salvar.nome.data
            flag_atividade = 'Qualidade'
            return app.redirect(url_for("download"))
        return app.redirect(url_for("qualidade"))
    #--------------------------------------------------------------------------------------------------------

    return render_template('qualidade.html',dados_dict = dados_dict,form_salvar=form_salvar,flag_ready=flag_ready)
#--------------------------------------------------------------------------------------------------------

@app.route('/eficiencia',methods=['GET','POST'])
def eficiencia():
    global nome_arquivo
    global flag_ready
    global results_dict
    form_salvar = FormSalvar()

    limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))

    #Procedimento para carregar arquivo
    if request.method == 'POST' and 'load_btn' in request.form:     
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            dados_dict = f_eficiencia.tratar_dados_eficiencia.carregar_dados(file = filename,folder=UPLOAD_FOLDER)
            # results_dict = f_eficiencia.tratar_dados_eficiencia.calculos(dados_dict=dados_dict)
            flag_ready = True
            flash('Dados carregados com sucesso',category='alert-success')
            return app.redirect(url_for('eficiencia'))
    #--------------------------------------------------------------------------------------------------------
        
    #Procedimento para salvar a planilha com as análises
    if 'salvar_btn' in request.form:            
        msg = f_eficiencia.tratar_dados_eficiencia.verificar_save(dados_dict=dados_dict)
        if msg != 'Arquivo salvo com sucesso': 
            flash(msg,category='alert-danger')
        else:
            limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))
            flag_atividade = 'Eficiencia'
            nome_arquivo = form_salvar.nome.data
            return app.redirect(url_for("download"))
        return app.redirect(url_for("eficiencia"))
    #--------------------------------------------------------------------------------------------------------
        
    return render_template('eficiencia.html')

@app.route('/',methods=['GET','POST'])
def vazamentos():
    global nome_arquivo
    global dados_empresa
    global flag_atividade
    global flag_registrado
    form_vazamentos = FormInfoVazamentos(data = {'empresa': dados_empresa['Empresa'],'cnpj':dados_empresa['CNPJ'],'endereco':dados_empresa['Endereço'],
                                                 'contato_nome':dados_empresa['Contato'],'contato_depto':dados_empresa['Departamento'],'contato_email':dados_empresa['E-mail'],
                                                 'contato_fone':dados_empresa['Telefone'],'e1':dados_empresa['Membro 1'],'e2':dados_empresa['Membro 2']
                                                 ,'e3':dados_empresa['Membro 3'],'e4':dados_empresa['Membro 4']})

    limpar_pasta(folder=os.path.join(app.root_path,UPLOAD_FOLDER))
    if 'reg_button' in request.form:
        dados_empresa = f_vazamentos.tratar_dados_vazamentos.tratar_dados(dados_empresa=dados_empresa,form_vazamentos=form_vazamentos)
        flag_registrado = True
        flash(f'Dados de atendimento registrados - {form_vazamentos.empresa.data}',category='alert-success')
        return app.redirect(url_for('vazamentos'))

    #Procedimento para carregar arquivo
    if request.method == 'POST' and 'load_btn' in request.form:  
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('Nenhum arquivo selecionado',category='alert-danger')
            return app.redirect(request.url)
        if file and allowed_file(file.filename):
            filename = file.filename

            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            f_vazamentos.tratar_dados_vazamentos.unzip(file=file,folder=os.path.join(app.root_path,UPLOAD_FOLDER))
            f_vazamentos.tratar_dados_vazamentos.relatorio(file=file.filename,folder=os.path.join(app.root_path,UPLOAD_FOLDER),dados_empresa=dados_empresa)
            nome_arquivo = f'{file.filename.split(".")[0]}/tabelas'
            flash('Dados carregados com sucesso',category='alert-success')
            flag_atividade = 'Vazamentos'
            print(nome_arquivo)
            return app.redirect(url_for("download"))
    #--------------------------------------------------------------------------------------------------------
        
    return render_template('vazamentos.html',form_vazamentos=form_vazamentos,flag_registrado=flag_registrado)
#--------------------------------------------------------------------------------------------------------

#Download de arquivos
@app.route('/download')
def download():
    global nome_arquivo
    global results_dict
    global dados_dict
    global flag_atividade
    nome = f'{nome_arquivo}.xlsx'
    if flag_atividade == 'Qualidade':
        f_qualidade.planilha_qualidade.criar_planilha(dados_dict=dados_dict,results_dict=results_dict,folder=os.path.join(app.root_path,UPLOAD_FOLDER),nome=nome)
    elif flag_atividade == 'Eficiencia':
        f_eficiencia.planilha_eficiencia.criar_planilha(dados_dict=dados_dict,results_dict=results_dict,folder=os.path.join(app.root_path,UPLOAD_FOLDER),nome=nome)
    elif flag_atividade == 'Vazamentos':
        nome = f'{nome_arquivo}.docx'
    uploads = os.path.join(app.root_path, app.config['UPLOAD_FOLDER'])
    return send_from_directory(directory=uploads, path=nome)
#--------------------------------------------------------------------------------------------------------

@app.route("/reset")
def reset():
    global dados_empresa
    global flag_registrado
    global flag_ready
    flag_ready = False
    flag_registrado = False
    dados_empresa = {
    'Empresa': '',
    'CNPJ': '',
    'Endereço': '',
    'Contato': '',
    'Departamento': '',
    'E-mail': '',
    'Telefone': '',
    'RT': '',
    'Membro 1': '',
    'Membro 2': '',
    'Membro 3': '',
    'Membro 4': '',
    }
    
    flash('Valores resetados',category='alert-success')
    return app.redirect(url_for('vazamentos'))
if __name__ == '__main__':
    app.run(debug=True)


