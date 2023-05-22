# url_pdf = pag_um.get_attribute('href')

# response = requests.get(url_pdf)
# with open('arquivo.pdf', 'wb') as arquivo:
#     arquivo.write(response.content)

# # Use o PyPDF2 para extrair o texto do PDF baixado
# with open('arquivo.pdf', 'rb') as arquivo_pdf:
#     leitor = PyPDF2.PdfReader(arquivo_pdf)

#     # Itera pelas páginas do PDF
#     for pagina_num in range(leitor.numPages):
#         pagina = leitor.getPage(pagina_num)
#         texto = pagina.extractText()
#         print(texto)
#         regex_cpf = r"\d{3}\.\d{3}\.\d{3}-\d{2}"
#         regex_cnpj = r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}"

#         # Extrai CPFs
#         cpfs = re.findall(regex_cpf, texto)

#         # Extrai CNPJs
#         cnpjs = re.findall(regex_cnpj, texto)

#         # Imprime os resultados
#         print("CPFs encontrados:", cpfs)
#         print("CNPJs encontrados:", cnpjs)