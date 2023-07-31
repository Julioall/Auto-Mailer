# Envio Automático de E-mails com Interface Gráfica

Este é um projeto em Python com uma interface gráfica (GUI) que permite o envio automático de e-mails para uma lista de destinatários a partir de um arquivo Excel.

## Descrição

O projeto é dividido em duas partes principais:

1. **Tela de Login** (`LoginScreen`): Esta tela exibe um formulário para o usuário inserir suas credenciais de e-mail (usuário e senha) e fazer o login. A autenticação é feita através de um servidor SMTP. Após a autenticação bem-sucedida, o usuário é redirecionado para a Tela Principal.

2. **Tela Principal** (`MainScreen`): Nesta tela, o usuário pode selecionar um arquivo Excel contendo uma lista de destinatários e enviar e-mails para eles. O arquivo Excel deve conter as colunas `colaborador`, `responsavel`, `versao`, `codigo` e `descricao`. Os e-mails são enviados através do servidor SMTP configurado na tela de login.

## Como Usar

1. Abra o aplicativo e insira suas credenciais de e-mail na tela de login.
2. Após fazer o login, você será redirecionado para a tela principal.
3. Clique no botão "Selecionar Arquivo" para escolher um arquivo Excel que contenha a lista de destinatários. Certifique-se de que o arquivo tenha as colunas necessárias (colaborador, responsavel, versao, codigo, descricao).
4. Após selecionar o arquivo, o caminho do arquivo será exibido na tela.
5. Clique no botão "Enviar" para enviar os e-mails para os destinatários da lista.
6. O progresso do envio será exibido na tela.

## Personalização

O projeto pode ser personalizado de acordo com as suas necessidades. Algumas possíveis melhorias incluem:

- Adicionar mais campos personalizados no arquivo Excel e utilizá-los na mensagem do e-mail.
- Personalizar a aparência da interface gráfica (cores, fontes, etc.).
- Adicionar mais recursos, como anexos aos e-mails ou a opção de enviar e-mails em horários programados.

## Requisitos

Certifique-se de ter as seguintes bibliotecas instaladas para executar o projeto:

- pandas
- smtplib
- ssl
- email
- tkinter

Você pode instalá-las usando o gerenciador de pacotes pip, por exemplo:

```powershell
pip install pandas
pip install secure-smtplib
pip install tk
```

## Observações

- O projeto foi desenvolvido considerando os provedores de e-mail Gmail, Outlook, Yahoo e um provedor personalizado (vetoquinol.com). Se desejar usar outro provedor, você pode adicionar um novo caso à função `autenticar` na classe `LoginScreen`.
- Certifique-se de permitir o acesso a aplicativos menos seguros (como aplicativos Python) na configuração da sua conta de e-mail, caso contrário, a autenticação pode falhar.
- O projeto foi criado para fins educacionais e pode ser aprimorado para ser mais seguro e eficiente em produção.

## Autor

Este projeto foi desenvolvido por [Julio Alves](https://github.com/Julioall).

Espero que este projeto seja útil para você! Divirta-se explorando e personalizando para suas necessidades. Se tiver alguma dúvida ou problema, sinta-se à vontade para entrar em contato.

***Bom desenvolvimento!***
