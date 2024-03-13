from flask_wtf import FlaskForm
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


class FormInfoVazamentos(FlaskForm):
    empresa = StringField('Nome da empresa', validators= [DataRequired()])
    cnpj = StringField('CNPJ da empresa', validators= [DataRequired()])
    endereco = StringField('Endereço da empresa (Logradouro; Bairro; Cidade - Estado)', validators= [DataRequired()])
    contato_nome = StringField('Nome da pessoa de contato na empresa', validators= [DataRequired()])
    contato_depto = StringField('Departamento da pessoa de contato', validators= [DataRequired()])
    contato_email = StringField('Email da pessoa de contato', validators= [DataRequired()])
    contato_fone = StringField('Número de telefone da pessoa de contato', validators= [DataRequired()])
    
    rt = StringField('Responsável Técnico', validators= [DataRequired()],default='Eng. Paulo Takao Okigami')
    e1 = SelectField('Equipe Técnica', validators= [DataRequired()],choices=['Esp. Haroldo Escudelario','Esp. Gidean de Sousa','Estg. Rodrigo Santana'])
    e2 = SelectField('Equipe Técnica',choices=['','Esp. Haroldo Escudelario','Esp. Gidean de Sousa','Estg. Rodrigo Santana'])
    e3 = SelectField('Equipe Técnica',choices=['','Esp. Haroldo Escudelario','Esp. Gidean de Sousa','Estg. Rodrigo Santana'])
    e4 = SelectField('Equipe Técnica',choices=['','Esp. Haroldo Escudelario','Esp. Gidean de Sousa','Estg. Rodrigo Santana'])

    setores = StringField('Setores (separar com vírgula)', validators= [DataRequired()])

    add_button = SubmitField('Registrar', validators= [DataRequired()])


class FormSalvar(FlaskForm):
    nome = StringField('Nome do arquivo', validators= [DataRequired()])
    salvar_btn = SubmitField('Salvar', validators= [DataRequired()])