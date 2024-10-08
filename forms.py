from flask_wtf import FlaskForm
from flask_wtf.form import _Auto
from wtforms import StringField,SubmitField,FloatField,IntegerField,SelectField
from wtforms.validators import DataRequired,Length,NumberRange,Regexp

class FormAddCarga(FlaskForm):
    nome_equip = StringField('Nome do equipamento', validators= [DataRequired()])
    potencia = FloatField('Potência (kW)', validators= [DataRequired()])
    dias = IntegerField('Nº de dias úteis', validators= [DataRequired(),NumberRange(min=1,max=31,message="Valor inválido")],default=22)
    fp = FloatField('Fator de potência', validators= [DataRequired()],default=1)
    fp_tipo = SelectField('Tipo de fator de potência', validators= [DataRequired()],choices=['Indutivo','Capacitivo'])
    qtd = IntegerField('Quantidade', validators= [DataRequired()],default=1)
    hr_inicio = StringField('Início', validators= [DataRequired(),Regexp(r"^[0-9][0-9][:][0-9][0-9]",message="O formato deve ser hh:mm")],default="00:00")
    hr_fim = StringField('Fim', validators= [DataRequired(),Regexp(r"^[0-9][0-9][:][0-9][0-9]",message="O formato deve ser hh:mm")],default="00:00")
    ponta = IntegerField('Início do horário de ponta', validators= [DataRequired(),NumberRange(min=0,max=21,message="Valor inválido")],default=17)

    add_button = SubmitField('Adicionar', validators= [DataRequired()])

file = open("arquivos/equipe_tecnica.txt", "r")
equipe0=[line.strip() for line in file.readlines()]
file.close()
equipe1 = equipe0.copy()
equipe1.insert(0,'')

class FormInfoVazamentos(FlaskForm):

    empresa = StringField('Nome da empresa', validators= [DataRequired()])
    cnpj = StringField('CNPJ da empresa', validators= [DataRequired()])
    endereco = StringField('Endereço da empresa (Logradouro; Bairro; Cidade - Estado)', validators= [DataRequired(),Regexp(r"^[:,.A-Za-záàâãéèêíïóôõöúçñÁÀÂÃÉÈÍÏÓÔÕÖÚÇÑ'\s0-9 °ºª\-]+[;][:.,A-Za-záàâãéèêíïóôõöúçñÁÀÂÃÉÈÍÏÓÔÕÖÚÇÑ'\s0-9 °ºª\-]+[;][:,.A-Za-záàâãéèêíïóôõöúçñÁÀÂÃÉÈÍÏÓÔÕÖÚÇÑ'\s0-9 °ºª\-]+$",message='Siga o formato de exemplo')])
    contato_nome = StringField('Nome da pessoa de contato na empresa', validators= [DataRequired()])
    contato_depto = StringField('Departamento da pessoa de contato', validators= [DataRequired()])
    contato_email = StringField('Email da pessoa de contato', validators= [DataRequired()])
    contato_fone = StringField('Número de telefone da pessoa de contato', validators= [DataRequired()])

    rt = StringField('Responsável Técnico', validators= [DataRequired()],default='Eng. Paulo Takao Okigami')
    e1 = SelectField('Equipe Técnica', validators= [DataRequired()],choices=equipe0)
    e2 = SelectField('Equipe Técnica',choices=equipe1)
    e3 = SelectField('Equipe Técnica',choices=equipe1)
    e4 = SelectField('Equipe Técnica',choices=equipe1)

    # setores = StringField('Setores (separar com vírgula)', validators= [DataRequired()])

    reg_button = SubmitField('Registrar', validators= [DataRequired()])

  

class FormSalvar(FlaskForm):
    nome = StringField('Nome do arquivo', validators= [DataRequired()])
    salvar_btn = SubmitField('Salvar', validators= [DataRequired()])