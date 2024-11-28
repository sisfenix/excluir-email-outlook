import win32com.client

# Conectando ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Obtendo a pasta padrão de e-mails (Inbox)
inbox = namespace.GetDefaultFolder(6)

# Função recursiva para percorrer as pastas e deletar e-mails
def deletar_emails(folder):
    print(f"\n\033[36mProcessando a pasta: {folder.Name} - Quantidade: {len(folder.Items)}\033[0m")
    # Filtrar e-mails para exclusão
    for item in folder.Items:
        try:
            if item.SenderEmailAddress in remetentes_para_deletar:
                print(f"[*** DELETAR ***] Deletando e-mail de: {item.SenderEmailAddress} | Assunto: {item.Subject}")
                item.Delete()
        except Exception as e:
            print(f"[ERRO] Não foi possível processar o item: {e}")
    
    # Processar subpastas
    for subfolder in folder.Folders:
        deletar_emails(subfolder)

# Array de remetentes cujos e-mails serão deletados
remetentes_para_deletar = [
    "remetente1@example.com",
    "remetente2@example.com"
]

# Iniciar o processo na pasta Inbox
if __name__ == "__main__":
    deletar_emails(inbox)
