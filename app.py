import openpyxl
from PIL import Image, ImageDraw, ImageFont

#conectar planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')

#conectar página planilha
sheet_alunos = workbook_alunos['Página1']

for indice, linha in enumerate( sheet_alunos.iter_rows(min_row=2)):
    nome = linha[0].value
    carga_horaria = linha[1].value
    dia = linha[2].value
    mes = linha[3].value
    ano = linha[4].value
    professor = linha [5].value


    #definindo fonte
    font_geral = ImageFont.truetype('./ariblk.ttf',40)
    font_assinatura = ImageFont.truetype('./ariali.ttf', 90)
    font_assinatura_prof = ImageFont.truetype('./ariali.ttf')

    imagem = Image.open('./certificado-padrao.png')
    desenhar = ImageDraw.Draw(imagem)

    desenhar.text((670,450), nome,fill='black',font=font_assinatura)
    desenhar.text((880,660),str(carga_horaria),fill='black',font=font_geral)
    desenhar.text((660,830),str(dia),fill='black',font=font_geral)
    desenhar.text((1000,830), mes,fill='black',font=font_geral)
    desenhar.text((1460,830),str(ano),fill='black',font=font_geral)

    imagem.save(f'./{nome} certificado.png')
