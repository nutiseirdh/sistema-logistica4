# Sistema de Logística 4.0 — gratuito e open-source

Aplicação inicial responsiva para uso local, construída com Python + Streamlit + SQLite.

## Funcionalidades incluídas
- Cadastro e consulta de materiais, SKU, unidade, estoque mínimo, localização e foto.
- Entradas e saídas com data/hora, motivo, responsável e documento opcional; valida saldo para evitar saída maior que o disponível.
- Cadastro de veículos, status e histórico de viagens com início e retorno.
- Cadastro de colaboradores, cargo, turno e perfil registrado.
- Dashboard com KPIs, gráficos e alerta de estoque baixo.
- Central de documentos com upload e download.
- Relatórios de estoque, movimentações, frota e viagens em Excel e PDF.

## Requisitos
- Python 3.10 ou superior
- Windows, Linux ou macOS
- Conexão à internet apenas para instalar as dependências na primeira execução; o uso posterior pode ser local.

## Instalação no Windows
1. Instale Python 3.10+ e marque a opção **Add Python to PATH**.
2. Extraia este pacote para uma pasta.
3. Abra o Prompt de Comando nessa pasta.
4. Execute:

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
streamlit run app.py
```

O navegador abrirá a aplicação. Se não abrir automaticamente, acesse `http://localhost:8501`.

## Backup
Feche a aplicação e copie a pasta `dados`, que contém o banco `logistica.db`, fotos e documentos anexados. Guarde cópias em local seguro.

## Observações importantes
- Este pacote é uma base funcional para protótipo/uso interno controlado, não uma implantação empresarial pronta para internet.
- O cadastro de permissões é apenas informativo: não implementa autenticação nem bloqueio por perfil. Antes de publicar na rede, implemente login seguro, autorização por função, trilha de auditoria, HTTPS, backups testados e política de retenção.
- A base SQLite foi escolhida para execução simples em um computador. Para múltiplos usuários simultâneos em rede, avalie PostgreSQL e uma API com autenticação.
- Os anexos ficam em disco local; defina limites de tamanho, antivírus e regras de acesso antes do uso produtivo.
- PDF e Excel são gerados sob demanda. Teste os relatórios com seus dados e procedimentos reais.
- Não inclua dados pessoais desnecessários e siga as políticas internas e a LGPD.


## Publicação web gratuita (teste)
1. Crie uma conta no GitHub e um repositório.
2. Envie `app.py` e `requirements.txt` para a raiz do repositório.
3. Acesse https://share.streamlit.io/ e entre com GitHub.
4. Escolha **Create app**, selecione o repositório, branch e `app.py`, e clique em **Deploy**.
5. Após a publicação, copie o endereço `*.streamlit.app`.

**Importante:** use somente dados fictícios na versão atual. A aplicação ainda não tem login nem autorização real, e o armazenamento SQLite/arquivos no disco da hospedagem gratuita não deve ser considerado persistente nem apropriado para operação institucional. Para uso real, implemente autenticação e banco/armazenamento persistentes antes de publicar dados. Repositório público também expõe o código-fonte.
