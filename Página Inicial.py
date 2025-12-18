import streamlit as st

# Configuração da página principal
st.set_page_config(
    page_title="Portal Financeiro",
    page_icon="🏢",
    layout="wide"
)

st.title("Portal de Ferramentas Contábeis")
st.markdown("---")

st.markdown("""
### Bem-vindo ao sistema centralizado.

Utilize o menu lateral à esquerda para navegar entre os módulos disponíveis:

- **Conciliador Bancário:** Ferramenta para cruzar dados do Extrato PDF com o Razão em Excel.
- **(Em breve) Novos Módulos:** Outras ferramentas serão adicionadas aqui.

---
**Status do Sistema:** ✅ Online
""")

# Dica: Se quiser que o login seja feito AQUI e valha para tudo,
# você pode mover a função check_password para cá no futuro.

