# Automação de Conformidade ISO 27001 (GRC)

## 📌 Visão Geral
Este projeto demonstra a aplicação de **Python e Inteligência Artificial** para otimizar processos críticos de Governança, Riscos e Conformidade (GRC). A solução automatiza a edição, versionamento e padronização de grandes volumes de documentos (100+) necessários para a certificação **ISO 27001**, reduzindo drasticamente o erro humano e o tempo de entrega.

---

## 🛡️ O Desafio (The Problem)
Na implementação da ISO 27001, a manutenção de políticas, normas e procedimentos exige revisões constantes. 
- **Escala:** Mais de 100 documentos .docx para edição simultânea.
- **Risco:** Alterações manuais "um a um" são propensas a erros de digitação, esquecimentos e inconsistências de versão.
- **Gargalo:** Mudanças solicitadas pela diretoria (ex: alteração de escopo, troca de logos ou nomes de responsáveis) levavam dias para serem replicadas em todo o SGSI (Sistema de Gestão de Segurança da Informação).

---

## 🚀 A Solução (The Solution)
Desenvolvi o **CyberDoc Automator Pro**, uma ferramenta em Python que utiliza bibliotecas de processamento de documentos para realizar operações em massa com segurança.

### Principais Funcionalidades:
1. **Mapeamento de Tags:** Varredura automática em pastas para identificar termos entre colchetes `[TAGS]` que precisam de preenchimento.
2. **Substituição Inteligente:** Troca de termos em parágrafos, tabelas, cabeçalhos e rodapés preservando a formatação original.
3. **Injeção de Mídia:** Substituição de tags de imagem por logos ou assinaturas digitais com redimensionamento automático.
4. **Versionamento Automático:** Atualização da tabela de "Histórico de Alterações" dentro de cada documento, inserindo data, versão, autor e descrição da mudança.
5. **Logs de Auditoria:** Geração de relatórios em Excel detalhando cada alteração realizada para fins de conformidade e revisão.

---

## 🛠️ Tecnologias Utilizadas
- **Linguagem:** Python 3.x
- **Bibliotecas:** `python-docx` (manipulação de Word), `openpyxl` (relatórios Excel), `tqdm` (barra de progresso), `re` (expressões regulares).
- **IA Generativa:** Utilizada como copiloto para acelerar o desenvolvimento do código e refinar a lógica de tratamento de erros.

---

## 📈 Impacto e Resultados
- **Eficiência:** Redução do tempo de edição de **~16 horas (2 dias de trabalho)** para **menos de 5 minutos**.
- **Qualidade:** Eliminação de 100% dos erros de "copia e cola" em termos padronizados.
- **Agilidade em GRC:** Capacidade de responder a mudanças regulatórias ou de diretoria em tempo real, mantendo a integridade do SGSI.

---

## 💡 Lições Aprendidas e Evolução
O projeto reforçou a importância da **revisão humana (Human-in-the-loop)**. Mesmo com a automação, o script gera um log detalhado para que o analista de cibersegurança valide se as alterações fazem sentido no contexto de cada política, garantindo que a automação sirva à segurança, e não o contrário.

---

> **Nota de Carreira:** Este projeto exemplifica como um profissional de Cibersegurança pode usar a tecnologia para escalar processos de governança, focando no que realmente importa: a análise estratégica de riscos.
