import keyboard
import pyautogui
import webbrowser
import time
from tkinter import *
import tkinter.font as tkFont
from tkinter.filedialog import askopenfilename
import pandas as pd
from datetime import *
import Integer

# Função para procurar e clicar no botão de enviar mensagem do WhatsApp
def moveclick(img, img2, img3, img4):
    images = (img,img2,img3,img4)
    find = False
    secs = 1
    while secs <= 1200:
        if keyboard.is_pressed('f8') == True:
            pyautogui.alert(text='Operação encerrada!', title='Aviso', button='OK')
            exit()
        for i in images:
            if pyautogui.locateOnScreen(i) != None:
                x, y = pyautogui.locateCenterOnScreen(i)
                pyautogui.click(x, y)
                find = True
                break
        if find == True:
            break
    time.sleep(.1)
    secs += 1

def open_file():
    # Abre arquivo com a base do push
    preview['text'] = 'Escolha o arquivo com a base de contatos\n'
    filepath = askopenfilename(title='Escolha a base de dados para o push',filetypes=[("Excel Files", "*.xlsx")])
    if not filepath:
        preview['text'] = 'Nenhum arquivo escolhido\n'
        return
    global base
    base = pd.read_excel(filepath)
    preview['text'] = f'{len(base)} Contatos carregados. Pronto para execução. Clique no método desejado\n'
    janela.title(f"Zappush - Base selecionada: {filepath.rsplit('/')[-1]}")

def nps():
    global nps, shift
    nps = 1
    shift = IntVar()

    info = Label(janela, text='Shift (1 = data da transação ontem, 2 = antiontem etc)', bg='#0B846D', fg='#D3D3D3', font=font3).pack()
    ent_shift = Entry(janela, textvariable=shift, width=5,font=font3).pack()
    botao_go = Button(janela, text='Mandar push de NPS', command=send, bg='#111B21', fg='#FFF', font=font2).pack(pady=20)

def send():
    global shift
    print(Integer.parseInt(shift))
    if shift == '':
        shift = 1
    # Carregar base all
    # df = pd.read_csv('C:/Users/Arthur/OneDrive - Consorciei Participações S.A/Documents/PowerBi/BASE ALL.csv',on_bad_lines='skip')
    df = pd.read_csv('BASE_ALL.csv', on_bad_lines='skip', sep=';')
    df = pd.DataFrame(df)
    df['Data da transação'] = pd.to_datetime(df['Data da transação'], format='%d/%m/%Y').dt.date

    # Remover linhas com nome do cliente duplicado
    df = df.drop_duplicates(subset='Signer1_Nome', keep='first')

    # Filtrar base pela data da transação
    df = df[df['Data da transação'] >= date.today() - timedelta(days=shift)]
    # api()

# Define se o método de push é via web ou api, de acordo com a escolha do usuário
def web():
    global plataforma
    plataforma = 'web'
    executar()

def api():
    global plataforma
    plataforma = 'api'
    executar()

# Executa a parada
def executar():
    try:
        # Prepara a mensagem a ser enviada para cada pessoa e abre o navegador com o link específico
        for row in range(len(base)):
            telefone = base.loc[row, 'Telefone']
            mensagens = []
            for n in range(1,7):
                try:
                    mensagens.append(base.loc[row, f'Mensagem {n}'])
                except:
                    pass
            set(mensagens)
            mensagens = [x for x in mensagens if str(x) != 'nan']

            for i in range(len(mensagens)):
                if pd.isna(mensagens[i]):
                    mensagens[i] = ""

            zap = f'https://{plataforma}.whatsapp.com/send?phone={telefone}&text={mensagens[0]}'
            webbrowser.open(zap)

            moveclick('assets/BTNsendwhite.png','assets/BTNsendblack.png',
                      'assets/BTNsendwhiteapi.png','assets/BTNsendblackapi.png')

            # Envia as mensagens 2 e 3 (opcionais)
            for i in range(1,len(mensagens)):
                if keyboard.is_pressed('f8') == True:
                    pyautogui.alert(text='Operação encerrada!', title='Aviso', button='OK')
                    exit()
                time.sleep(.3)
                pyautogui.write(mensagens[i])
                time.sleep(.3)
                pyautogui.press('enter')

            # Enquanto a mensagem não tiver sido enviada, aguardar
            while pyautogui.locateOnScreen('assets/sendingmsgwhite.png')!=None or \
                    pyautogui.locateOnScreen('assets/sendingmsgblack.png')!=None:
                if keyboard.is_pressed('f8') == True:
                    pyautogui.alert(text='Operação encerrada!', title='Aviso', button='OK')
                    exit()
                time.sleep(.1)

            # Fechar navegador
            time.sleep(.2)
            if plataforma == 'api':
                pyautogui.hotkey('win', 'down')
                time.sleep(.3)
            pyautogui.hotkey('ctrl','w')
    except NameError:
        pyautogui.alert('Nenhuma base selecionada. Clique em "Abrir arquivo"', title='Aviso', button='OK')

def btninstrucoes():
    if instrucoes["text"] == "":
        instrucoes["text"] = '1 - Prepare uma base no Excel com colunas "Telefone" e "Mensagem 1"\n' \
                             '2 - Opcionalmente, adicione colunas "Mensagem 2", "Mensagem 3" etc\n' \
                             '(Essa é a ordem das mensagens a serem enviadas a cada pessoa)\n' \
                             '3 - Salve a planilha, clique em "Abrir arquivo" e depois em algum método\n\n' \
                             '-> O programa será executado até o fim da lista\n' \
                             '-> Para encerrar a operação, pressione a tecla F8\n'
    else:
        instrucoes["text"] = ""

def btncontato():
    help = 'Oi, Arthur. Quero falar sobre o Bot do WhatsApp'
    webbrowser.open(f'https://web.whatsapp.com/send?phone=5511954191629&text={help}')


# Cria a janela do front
janela = Tk()
janela.title('Zappush')
janela.configure(bg='#0B846D')

font1 = tkFont.Font(family="Arial Black", size=18, weight="bold")
font2 = tkFont.Font(family="Arial Black", size=12, weight="bold")
font3 = tkFont.Font(family="Arial Black", size=9)

frame = Frame(master=janela,bg='#111B21')
frame.pack(padx=10,pady=10,fill=BOTH,expand=TRUE)

frame.grid_columnconfigure([0,1], weight=1)

texto1 = Label(frame, text='Bot para envio de push via WhatsApp',bg='#111B21',fg='#D3D3D3',font=font1).grid(columnspan=2,padx=80,pady=20)
# texto2 = Label(frame, text='Selecione um método para mandar o push!',bg='#111B21',fg='#D3D3D3',font=font2).grid(columnspan=2,padx=5,pady=5)
botao_web = Button(frame,text='WEB (Navegador)',command=web,bg='#25B43C',fg='#FFF',font=font2).grid(row=2,column=0,padx=5,pady=5,sticky=E)
botao_api = Button(frame,text='API (Aplicativo)', command=api,bg='#25B43C',fg='#FFF',font=font2).grid(row=2,column=1,padx=5,pady=5,sticky=W)
botao_instrucoes = Button(frame,text='Instruções',bg='#D3D3D3',command=btninstrucoes,font=font2).grid(row=3,column=0,padx=5,pady=5,sticky=E)
botao_abrir_arquivo = Button(frame,text='Abrir arquivo',bg='#D3D3D3',command=open_file,font=font2).grid(row=3,column=1,padx=5,pady=5,sticky=W)


instrucoes = Label(frame, text="",bg='#111B21',fg='#D3D3D3',font=font2)
instrucoes.grid(columnspan=2,padx=5,pady=5)

preview = Label(frame, text="",bg='#111B21',fg='#D3D3D3',font=font3)
preview.grid(columnspan=2,padx=5,pady=5)


botao_nps = Button(janela,text='NPS',bg='#D3D3D3',command=nps,font=font3)
botao_nps.pack(side=RIGHT,padx=20,pady=5)
contato = Button(janela, text="Dúvida ou sugestão? Clique aqui",bg='#0B846D',fg='#D3D3D3',font=font3,borderwidth=0,command=btncontato)
contato.pack()



janela.mainloop()
