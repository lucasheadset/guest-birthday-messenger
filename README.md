# 🚀 Guest Messenger // Mensageiro de Aniversariantes

![Demonstração da Interface]((https://imgur.com/a/95QtXLL))

## 📄 Descrição

O Mensageiro de Aniversariante é uma aplicação de desktop desenvolvida em Python para automatizar o envio de mensagens de aniversário para hóspedes de hotel via WhatsApp. O projeto foi criado como uma solução prática para um problema de negócio real, com o objetivo de otimizar o tempo da equipe e melhorar a estratégia de retenção e marketing de relacionamento com o cliente.

Inicialmente o projeto foi criado para atender um hotel no qual eu trabalhava, porém, após minha saída do mesmo, o projeto foi descontinuado e não há intenção de finalização.

## ✨ Funcionalidades Principais

* **Busca Automática:** Identifica os aniversariantes do dia a partir de uma base de dados.
* **Interface Gráfica (GUI):** Desenvolvida com a biblioteca Tkinter, oferecendo uma experiência de usuário amigável e intuitiva.
* **Envio em Segundo Plano:** O envio de mensagens em massa é executado em uma thread separada para não congelar a interface do programa.
* **Busca Avançada:** Permite pesquisar hóspedes por Nome, CPF, Telefone ou Data de Nascimento, com um seletor de DDI internacional.
* **Sistema de Ajuda:** Inclui um chatbot de ajuda desenvolvido com HTML, CSS e JavaScript para guiar o usuário.
* **Gerenciamento de Exceções:** Permite adicionar hóspedes a uma lista para não receberem a mensagem.

## ❌ Funcionalidades Não-Finalizadas

* **Modo escuro:** Com a intenção de mudar o tema do app, função está atualmente bugada, causando erro gráfico na interface.
* **Notificações** Apesar de haver notificações na abertura do app, inicialmente a ideia era que na versão final, também tivéssemos uma notificação quando finalizado o envio. Atualmente, tal função não está presente no projeto.
* **Configurações no geral** Apesar da tela de configurações ser funcional, as configurações feitas nela não fazem efeito no programa em si.
* **Iniciar envio por barra de tarefas** Apesar de haver a opção o programa só irá crashar.
* **Modo VIP** Inicialmente com a intenção de ser uma forma de diferenciar hóspedes habitues, o mesmo não se encontra finalizado nessa versão.

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** Python
* **Interface Gráfica:** Tkinter
* **Manipulação de Dados:** Pandas
* **Automação WhatsApp:** PyWhatKit
* **Desenvolvimento Assistido:** Ferramentas de IA Generativa para aceleração de código e depuração.

## 🚀 Como Rodar o Projeto

1.  Clone este repositório.
2.  Instale as dependências necessárias:
    `pip install -r requirements.txt`
3.  Execute o script principal:
    `python beta.py`

## 🧠 O Que Eu Aprendi com Este Projeto

* Desenvolvimento de interfaces gráficas funcionais com a biblioteca Tkinter.
* Manipulação e filtragem de dados de forma eficiente utilizando Pandas.
* A importância de desacoplar a lógica da aplicação da sua fonte de dados.
* Implementação de multithreading para executar tarefas longas sem impactar a experiência do usuário.

