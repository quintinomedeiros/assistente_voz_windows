import speech_recognition as sr
import os
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

# Função para ouvir e reconhecer a fala
def ouvir_microfone():
    # Habilita o microfone do usuário
    microfone = sr.Recognizer()
    mic_list = sr.Microphone.list_microphone_names()

    if not mic_list:
        messagebox.showerror("Erro", "Nenhum microfone encontrado. Verifique a conexão e as permissões.")
        return

    # Usando o microfone, tenta o primeiro microfone disponível
    with sr.Microphone(device_index=0) as source:
        # Reduz ruídos de fundo
        microfone.adjust_for_ambient_noise(source)
        label_status.config(text="Aguardando comando...")
        root.update()  # Atualiza a interface para mostrar a mensagem

        # Armazena o que foi dito numa variável
        audio = microfone.listen(source)

    try:
        # Passa a variável para o algoritmo reconhecedor de padrões
        frase = microfone.recognize_google(audio, language='pt-BR')

        # Executa comandos baseados na frase reconhecida
        if "navegador" in frase:
            os.system('start Chrome.exe')
            messagebox.showinfo("Ação", "Abrindo navegador...")

        elif "Excel" in frase:
            os.system('start Excel.exe')
            messagebox.showinfo("Ação", "Abrindo Excel...")

        elif "PowerPoint" in frase:
            os.system('start POWERPNT.exe')
            messagebox.showinfo("Ação", "Abrindo PowerPoint...")

        elif "pasta" in frase or "projetos" in frase:
            os.system(r'start explorer "Y:\!COEAD\1. Projetos\1. Direito Constitucional\Direito Constitucional (20 e 60horas)"')
            messagebox.showinfo("Ação", "Abrindo a pasta de Projetos...")

        elif "Edge" in frase or "Edi" in frase:
            os.system('start msedge.exe')
            messagebox.showinfo("Ação", "Abrindo Edge...")

        elif "Fechar" in frase:
            messagebox.showinfo("Ação", "Fechando assistente...")
            root.quit()

        else:
            messagebox.showinfo("Comando não reconhecido", f"Você disse: {frase}")

    except sr.UnknownValueError:
        messagebox.showerror("Erro", "Não entendi o que você disse")
    finally:
        # Reseta a mensagem para o estado original
        label_status.config(text="Clique no botão para falar")

# Configuração da interface gráfica
def iniciar_assistente():
    global root, label_status
    root = tk.Tk()
    root.title("Assistente Virtual")
    
    # Carrega a imagem do robô
    bg_image = Image.open("robo.png")
    bg_photo = ImageTk.PhotoImage(bg_image)

    # Define o tamanho da janela com base na imagem
    root.geometry(f"{bg_image.width}x{bg_image.height}")

    # Cria um label para a imagem de fundo
    bg_label = tk.Label(root, image=bg_photo)
    bg_label.place(x=0, y=0, relwidth=1, relheight=1)

    # Cria um label para mostrar o status
    label_status = tk.Label(root, text="Clique para falar", font=("Arial", 11), bg='black', fg='white')
    label_status.place(x=160, y=140)  # Posiciona no "corpo" do robô

    # Botão "Ouvir"
    botao_ouvir = tk.Button(root, text="Ouvir", command=ouvir_microfone, font=("Arial", 12, "bold"), bg='#106DB3', fg='yellow')
    botao_ouvir.place(x=130, y=230)  # Posiciona no "corpo" do robô

    root.mainloop()

# Inicia o assistente
if __name__ == "__main__":
    iniciar_assistente()
