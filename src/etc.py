#!/usr/bin/env python
import pandas
import sqlite3

# Connects and collects data from the database and transforms it into a dataframe
connection = sqlite3.connect('database.db')
query_string = "SELECT * FROM logsmodel;"
dataframe = pandas.read_sql_query(query_string,connection)
# Transform column value into datetime for date operations
dataframe['received_at'] = pandas.to_datetime(dataframe['received_at'])
dataframe['create_at'] = pandas.to_datetime(dataframe['create_at'])
# Replaces redundant values ​​to facilitate calculations
dataframe['aplicacao'].replace('debito', 'fatura', inplace=True)
dataframe['aplicacao'].replace('roteiro', 'leiturista', inplace=True)
dataframe['aplicacao'].replace('contato', 'telefone', inplace=True)
# Remove weekend from dataframe
dataframe = dataframe[~dataframe['create_at'].dt.weekday.isin([5,6])]
# Calculates query count
dataframe['created_date'] = dataframe['create_at'].dt.date
quantidade_de_dias = dataframe['created_date'].nunique()
solicitacoes_total = dataframe['informacao'].count()
solicitacoes_media = int(solicitacoes_total / quantidade_de_dias)
faturas_total = dataframe[dataframe['aplicacao'] == 'fatura']['informacao'].count()
faturas_media = int(faturas_total / quantidade_de_dias)
leiturista_total = dataframe[dataframe['aplicacao'] == 'leiturista']['informacao'].count()
leiturista_media = int(leiturista_total / quantidade_de_dias)
telefone_total = dataframe[dataframe['aplicacao'] == 'telefone']['informacao'].count()
telefone_media = int(telefone_total / quantidade_de_dias)
agrupamento_total = dataframe[dataframe['aplicacao'] == 'agrupamento']['informacao'].count()
agrupamento_media = int(agrupamento_total / quantidade_de_dias)
informacao_total = dataframe[dataframe['aplicacao'] == 'medidor']['informacao'].count()
informacao_media = int(informacao_total / quantidade_de_dias)
historico_total = dataframe[dataframe['aplicacao'] == 'historico']['informacao'].count()
historico_media = int(historico_total / quantidade_de_dias)
outros_total = solicitacoes_total - (faturas_total + leiturista_total + telefone_total + agrupamento_total + informacao_total + historico_total)
outros_media = int(outros_total / quantidade_de_dias)
porcentagem_sucesso = int((dataframe[dataframe['is_sucess'] == 1]['informacao'].count() / solicitacoes_total) * 100)
dataframe.dropna(subset='received_at', inplace=True)
dataframe['duracao'] = dataframe['create_at'] - dataframe['received_at']
tempo_medio = dataframe['duracao'].mean()
# Write the result of the counts
print("**Relatório de utilização do chatbot:**")
print("")
print(f"Contagem de solicitações:\nTotal: {solicitacoes_total}, Média: {solicitacoes_media}")
print(f"Fatura pendentes em PDF:\nTotal: {faturas_total}, Média: {faturas_media}")
print(f"Roteiro de leiturista:\nTotal: {leiturista_total}, Média: {leiturista_media}")
print(f"Telefone do cliente:\nTotal: {telefone_total}, Média: {telefone_media}")
print(f"Análise de agrupamento:\nTotal: {agrupamento_total}, Média: {agrupamento_media}")
print(f"Informações do medidor\nTotal: {informacao_total}, Média: {informacao_media}")
print(f"Histórico de notas:\nTotal: {historico_total}, Média: {historico_media}")
print(f"Outras solicitações:\nTotal: {outros_total}, Média: {outros_media}")
print("")
print(f"O chatbot está sendo utilizado a {quantidade_de_dias} dias, atendendo as solicitações num tempo médio de {tempo_medio.seconds} segundos, com precisão de {porcentagem_sucesso}% de sucesso!") # type: ignore
