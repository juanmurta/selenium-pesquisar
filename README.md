# Web Scraper de Ofertas (Google Shopping & Buscapé)

## Descrição
Este projeto utiliza **Selenium** e **Pandas** para buscar ofertas de produtos no **Google Shopping** e **Buscapé**. Ele permite filtrar os resultados com base em termos proibidos e uma faixa de preço desejada. Ao final, os resultados são exportados para um arquivo Excel e enviados por e-mail via **Outlook**.

## Funcionalidades Principais
- **Leitura de uma lista de produtos** a partir de um arquivo Excel (`buscas.xlsx`).
- **Busca automatizada no Google Shopping e Buscapé** utilizando **Selenium**.
- **Filtragem de produtos**:
  - Exclui produtos que contenham termos banidos.
  - Garante que todos os termos do nome do produto estejam presentes na oferta.
  - Considera apenas ofertas dentro da faixa de preço especificada.
- **Exportação das ofertas encontradas** para um arquivo `Ofertas.xlsx`.
- **Envio automático de e-mail via Outlook** com a lista de ofertas.

## Como Usar a Ferramenta
1. Instale as dependências necessárias utilizando `pip install selenium pandas pywin32`.
2. Configure o WebDriver do Chrome (Baixe o ChromeDriver correspondente à sua versão do navegador em https://chromedriver.chromium.org/downloads).
3. Crie um arquivo `buscas.xlsx` contendo:
   - **Nome**: Nome do produto a ser pesquisado.
   - **Termos banidos**: Palavras que devem ser evitadas na busca.
   - **Preço mínimo**: Valor mínimo da oferta.
   - **Preço máximo**: Valor máximo da oferta.
4. Execute o script normalmente (`python scraper.py`).
5. O programa irá:
   - Realizar as buscas no **Google Shopping** e **Buscapé**.
   - Filtrar os produtos de acordo com os critérios definidos no Excel.
   - Exportar os resultados para `Ofertas.xlsx`.
   - Enviar um e-mail com os produtos encontrados.

## Como Contribuir
1. Faça um fork do repositório.
2. Crie um branch para suas alterações: `git checkout -b minha-nova-funcionalidade`
3. Commit suas mudanças: `git commit -m "Adiciona nova funcionalidade"`
4. Faça push para o branch: `git push origin minha-nova-funcionalidade`
5. Abra um Pull Request.

Para mais detalhes sobre como contribuir, veja a [documentação oficial do GitHub sobre Pull Requests](https://docs.github.com/pt/pull-requests/collaborating-with-pull-requests).
