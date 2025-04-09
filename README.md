# PDFCropperPro ✂️📄

Ferramenta Python para recortar, redimensionar e exportar páginas de PDFs com alta qualidade (400 DPI). Desenvolvida com uma interface gráfica interativa, permite manipular PDFs de forma prática e eficiente, ideal para ajustes personalizados e impressão de alta resolução.

## 📋 Descrição

O **PDFCropperPro** é uma aplicação Python que utiliza a biblioteca `PyMuPDF` (`fitz`) e `Tkinter` para oferecer uma solução completa de manipulação de PDFs. Com ele, você pode recortar áreas específicas de páginas, aplicar recortes a múltiplas páginas, rotacionar páginas, ajustar o zoom, e exportar ou salvar os resultados em alta qualidade (400 DPI). O projeto foi otimizado para garantir uma experiência fluida e resultados profissionais.

### Funcionalidades
- **Abrir e Visualizar PDFs**: Carregue arquivos PDF e navegue entre as páginas com facilidade.
- **Recorte Personalizado**: Selecione áreas específicas de uma página para recortar, com visualização em tempo real.
- **Aplicar Recortes a Múltiplas Páginas**: Aplique o mesmo recorte a todas as páginas ou a páginas selecionadas.
- **Exportar Páginas**: Exporte páginas específicas para um novo PDF com recortes aplicados.
- **Alta Resolução (400 DPI)**: Salve PDFs com qualidade otimizada para impressão, ajustando o DPI para 400.
- **Rotação de Páginas**: Rotacione páginas em incrementos de 90° para facilitar a visualização e edição.
- **Zoom e Navegação**: Ajuste o zoom com controles deslizantes ou roda do mouse e navegue entre páginas com botões ou teclas de seta.
- **Salvar e Carregar Configurações**: Salve suas configurações de recorte em um arquivo JSON e carregue-as posteriormente.
- **Desfazer e Limpar Recortes**: Redefina ou remova recortes aplicados com facilidade.

## 📦 Pré-requisitos

- **Python 3.8 ou superior**
- Bibliotecas necessárias:
  - `PyMuPDF`: Para manipulação de PDFs.
  - `Pillow`: Para processamento de imagens.
  - `tkinter`: Para a interface gráfica (geralmente incluído com o Python).

Instale as dependências com o seguinte comando:
```bash
pip install PyMuPDF Pillow
xecute o script:

bash
Copy
python Corta_pdf_Ajusta_Tamanho_do_conteudo.py
📝 Funcionalidades
Redimensiona páginas de PDFs para conteúdo específico.

Remove margens desnecessárias.

🤝 Contribuição
Contribuições são bem-vindas! Siga estes passos:

Faça um "fork" do projeto.

Crie uma branch (git checkout -b feature/nova-funcionalidade).

Faça commit das mudanças (git commit -m 'Adiciona nova funcionalidade').

Envie para o repositório (git push origin feature/nova-funcionalidade).
