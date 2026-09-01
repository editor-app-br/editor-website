# AVISO LEGAL E DE LICENCIAMENTO (LEGAL NOTICE & ATTRIBUTION)

**Projeto:** Editor (editor.app.br)  
**Licença Principal:** GNU Affero General Public License Version 3 (GNU AGPL v3)  
**Patrocinador & Mantenedor:** PRATA CONSULTORIA EM TECNOLOGIA DA INFORMACAO LTDA  
**CNPJ:** 48.889.573/0001-44  
**Site do Patrocinador:** [https://prata.dev.br](https://prata.dev.br)  
**Contato Institucional:** contato@prata.dev.br  

---

## 1. Declaração de Software Livre e AGPL v3 (Seção 13)

Este software é um Software Livre distribuído sob os termos da **GNU Affero General Public License v3.0**.

De acordo com a **Seção 13** da licença GNU AGPL v3 (*Remote Network Interaction*), todos os usuários que interagem com este software através de uma rede de computadores têm o direito expresso e irrestrito de obter o Código Fonte Correspondente (*Corresponding Source Code*).

- **Página de código-fonte:** https://editor.app.br/source
- **Aviso legal:** https://editor.app.br/legal
- **Repositório público:** https://github.com/editor-app-br/editor-website
- **Contato (código-fonte):** contato@prata.dev.br

---

## 2. Atribuições e Licenças de Terceiros (Upstream Components)

### 2.1. ONLYOFFICE DocumentServer
- **Autor / Detentor dos Direitos:** © Ascensio System SIA
- **Licença:** GNU Affero General Public License v3 (AGPL-3.0) com Cláusulas Adicionais de Marca e UI.
- **Aviso de Marca:** ONLYOFFICE™ é uma marca registrada de Ascensio System SIA. Este projeto utiliza o motor de código aberto ONLYOFFICE Community Edition em estrita conformidade com as regras comunitárias de exibição de logotipo e créditos na interface de edição.
- **Website:** [https://www.onlyoffice.com](https://www.onlyoffice.com)

### 2.2. ZIZIYI office-website
- **Autor original:** © baotlake (ZIZIYI)
- **Licença:** GNU AGPL v3
- **Contribuição:** Arquitetura de cliente local, emulação de socket no navegador e orquestração de WebAssembly.
- **Repositório:** [https://github.com/baotlake/office-website](https://github.com/baotlake/office-website)

### 2.3. CryptPad onlyoffice-x2t-wasm
- **Autor original:** XWiki SAS / Equipe CryptPad
- **Licença:** GNU AGPL v3
- **Contribuição:** Compilação do motor de conversão de documentos C++ x2t para WebAssembly (WASM).

---

## 3. Modelo de Integração e Isolamento (Thin-Embed Boundary)

A integração do editor.app.br com sistemas web externos e corporativos ocorre através do padrão de isolamento **Thin Embed (<iframe> + postMessage)**:

1. **Separação de Processo e Domínio:** O editor é executado sob sua própria origem isolada.
2. **Nenhuma Obra Derivada no Host:** O sistema host apenas transmite e recebe streams de bytes (ArrayBuffer) de documentos criados pelo usuário, sem incorporar ou derivar do código AGPL.
3. **Privacidade e LGPD:** Todos os documentos são processados localmente na memória RAM do navegador do usuário final. O servidor editor.app.br não armazena, não inspeciona e não mantém cópia de nenhum documento de usuário.
