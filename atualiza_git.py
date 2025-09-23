import os
import subprocess
from datetime import datetime

def rodar_comando(comando):
    resultado = subprocess.run(comando, shell=True, capture_output=True, text=True)
    if resultado.returncode == 0:
        print(f"✔ Sucesso: {comando}")
    else:
        print(f"❌ Erro no comando: {comando}")
        print(resultado.stderr)

if __name__ == "__main__":
    # 1. Rodar o script de gerar HTML
    print("🔄 Gerando relatório HTML...")
    rodar_comando("python gerar_html.py")

    # 2. Adicionar mudanças ao Git
    print("📂 Adicionando arquivos modificados...")
    rodar_comando("git add .")

    # 3. Criar commit com data/hora
    mensagem = f'Atualização automática {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}'
    print("📝 Criando commit...")
    rodar_comando(f'git commit -m "{mensagem}"')

    # 4. Enviar para o GitHub
    print("🚀 Enviando para o GitHub...")
    # troque 'main' por 'master' se o seu repositório estiver usando master
    rodar_comando("git push origin master")

    print("✅ Relatório atualizado no GitHub Pages!")
