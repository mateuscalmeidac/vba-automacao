# 📧 envio automático de e-mails de contadores de crachás

este projeto contém um script em **vba** que automatiza o envio de e-mails via **outlook**, a partir de informações armazenadas em uma planilha do **excel**.  
o objetivo é facilitar a solicitação mensal dos contadores de impressoras de crachás.

---

## 🚀 funcionalidades
- lê os dados (nome, e-mail, série, modelo e centro de custo) diretamente de uma planilha do excel.  
- envia e-mails personalizados para cada destinatário via outlook.  
- adiciona no corpo do e-mail os detalhes da impressora vinculada.  
- inclui o mês/ano atual automaticamente no assunto do e-mail.  
- insere mensagem padrão de agradecimento e aviso para desconsiderar caso já tenha enviado.  

---

## 📂 estrutura esperada da planilha
a macro considera que os dados estejam organizados a partir da linha 2, com os seguintes campos:

| coluna | informação         |
|--------|--------------------|
| a      | nome               |
| b      | e-mail             |
| c      | série              |
| d      | modelo             |
| e      | centro de custo    |

---

## ▶️ como usar
1. abra o arquivo excel que contém os dados na primeira aba.  
2. insira o código vba no editor (atalho `alt + f11`).  
3. ajuste os campos da planilha caso necessário.  
4. execute a macro `EnviarEmailsCrachas`.  
5. todos os e-mails serão enviados automaticamente pelo outlook.  

---

## ⚠️ requisitos
- microsoft excel  
- microsoft outlook configurado com conta de e-mail ativa  
- macros habilitadas no excel  

---

## ✍️ autor
desenvolvido por **matheus cruz - ti**
            .Body = corpoEmail
            .Send
        End With
    Next i

    MsgBox "Todos os e-mails foram enviados com sucesso!", vbInformation
End Sub
