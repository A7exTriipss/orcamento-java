# orcamento-java
Aplicação em Java para criação de orçamentos, preenchendo automaticamente um modelo Excel com dados de clientes e serviços. Feito em Java puro com Apache POI.



# Projeto Orçamento em Java (Apache POI)

Este é um projeto em **Java puro** que utiliza a biblioteca **Apache POI** para manipular arquivos Excel (.xlsx).

---

## Estrutura do projeto

orcamento-java/
├─ src/ # Código .java
├─ libs/ # Dependências (ver instruções abaixo)
├─ modelo_orcamento.xlsx
└─ README.md



## Dependências

Para compilar e rodar, você precisa baixar os seguintes **jars** e colocá-los dentro da pasta `libs/`:

- poi-5.2.5.jar
- poi-ooxml-5.2.5.jar
- poi-ooxml-schemas-5.2.5.jar
- xmlbeans-5.1.1.jar
- commons-collections4-4.4.jar
- commons-compress-1.21.jar
- commons-io-2.11.0.jar
- curvesapi-1.06.jar (opcional, para gráficos avançados)

> Todos podem ser baixados em [Apache POI](https://poi.apache.org/download.html).

---

## Como compilar e rodar

### Linux / Mac
```bash
javac -cp "libs/*" -d bin src/*.java
java -cp "bin:libs/*" PreenchedorExcelModelo



Fluxo Git diário

Puxar atualizações:

git pull origin main


Adicionar alterações e commit:

git add .
git commit -m "Descrição do que foi feito"
git push origin main


---

##✅ Benefícios desta abordagem:##
- Projeto no GitHub **mais leve** (sem jars grandes)  
- Qualquer PC consegue rodar seguindo instruções do README  
- Permite versionar apenas o **código-fonte**, modelo Excel e scripts  

---

Se quiser, posso te criar **uma versão final do `.gitignore` + README` + scripts (`run.sh` e `run.bat`)** prontos para você subir **agora mesmo no GitHub**, sem os jars, já organizados pro seu estudo.  

Quer que eu faça isso?
