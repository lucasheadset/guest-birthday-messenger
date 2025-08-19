# 🚀 Guest Messenger // Mensageiro de Aniversariantes

![Demonstração da Interface](https://i.imgur.com/yPCr7df.png)

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
  
## 🔮 Próximos Passos e Melhorias Futuras

Embora o projeto seja totalmente funcional em seu estado atual, algumas melhorias futuras poderiam ser implementadas para torná-lo ainda mais robusto:

* **Implementação de Tema Escuro (Dark Mode):** Adicionar um tema alternativo para melhorar a usabilidade em ambientes com pouca luz.
* **Validação de Contatos com Selenium:** Desenvolver um módulo que utiliza Selenium para verificar a validade de números de WhatsApp antes do envio, aumentando a taxa de sucesso.
* **Integração com LLMs:** Conectar o chatbot de ajuda a uma API de um modelo de linguagem (como a do ChatGPT) para fornecer respostas mais dinâmicas e inteligentes.
* **Notificações Avançadas:** Criar um sistema de notificações para avisar o usuário sobre a conclusão dos envios e possíveis falhas.
* **Sistema de VIPs:** Finalizar a funcionalidade de "VIPs" para permitir o envio de mensagens diferenciadas para clientes recorrentes.
 
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





