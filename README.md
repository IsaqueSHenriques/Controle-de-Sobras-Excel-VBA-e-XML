# Controle-de-Sobras-Excel-VBA-e-XML
Automação em VBA para importar sobras de planos de corte em XML e organizá-las por material no Excel.

Importação de Sobras de Corte com Excel e VBA

Antes desta automação, as sobras geradas no software de plano de corte eram copiadas manualmente para uma planilha. Era necessário consultar cada sobra no software e registrar suas informações na aba do material correspondente.

O software já permitia exportar um arquivo XML com esses dados. Desenvolvi uma macro em VBA que lê esse arquivo e preenche a planilha automaticamente.

Como funciona

1-O usuário informa a data de corte e seleciona um ou mais arquivos XML.
2-Macro lê as xml's e identifica as sobras

3-Ela registra lote, data, comprimento, largura, quantidade e descrição na aba correspondente ao material.

Tecnologias utilizadas

Excel, VBA, XML e MSXML 6.0.

Objetivo

Reduzir o trabalho de digitação e facilitar a organização das sobras por material.
