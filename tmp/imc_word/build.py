from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
from copy import deepcopy
import hashlib
ref=Path(r'C:\Users\Instrutor\Downloads\Respostas_Aula11_Computacao_em_Nuvem.docx')
out=Path('tmp/imc_word/Aula10Devops-IMC.docx')
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
w='{'+ns['w']+'}'
z=ZipFile(ref); root=E.fromstring(z.read('word/document.xml')); body=root.find('w:body',ns)
ps=body.findall('w:p',ns); table=deepcopy(body.find('w:tbl',ns)); sect=deepcopy(body.find('w:sectPr',ns))
Path('tmp/imc_word/artifact.md').write_text('Reference: '+str(ref)+'\nSHA256: '+hashlib.sha256(ref.read_bytes()).hexdigest()+'\nA4 portrait 11906 x 16838 twips; margins top/bottom 1134, left/right 1247. Preserve styles, section and package parts. Clone title p0, heading p4, body p5, caption p12; screenshot tables one cell with grid border and gray placeholder. Remove prior drawings from cloned cells. Five pages: overview, tests, failure, correction, answers. Keep author and RA from reference. Screenshots remain blank. Only document.xml and core metadata editable.\n',encoding='utf8')
for child in list(body): body.remove(child)
def p(text,kind=5,page=False):
    a=deepcopy(ps[kind]); prop=a.find('w:pPr',ns)
    for x in list(a):
        if x is not prop: a.remove(x)
    if page:
        if prop is None: prop=E.SubElement(a,w+'pPr'); a.insert(0,prop)
        E.SubElement(prop,w+'pageBreakBefore')
    r=E.SubElement(a,w+'r'); t=E.SubElement(r,w+'t'); t.text=text
    body.append(a); return a
def box():
    t=deepcopy(table)
    for cell in t.findall('.//w:tc',ns):
        for x in list(cell):
            if x.tag != w+'tcPr': cell.remove(x)
        para=E.SubElement(cell,w+'p'); run=E.SubElement(para,w+'r'); rp=E.SubElement(run,w+'rPr'); E.SubElement(rp,w+'color').set(w+'val','787878'); E.SubElement(run,w+'t').text='Espaço reservado para o print'
    h=t.find('.//w:trHeight',ns); h.set(w+'val','3900'); h.set(w+'hRule','atLeast')
    trpr=t.find('.//w:trPr',ns); E.SubElement(trpr,w+'cantSplit')
    body.append(t); p('',14)
p('Testes do CalcIMC no GitHub Actions',0)
p('DevOps | Aula 10 | Exercício 1',1)
p('Nome: Raphael Gouvêa Reche   RA: 26164063',2)
p('Professor: Daniel Vieira   |   Data: ____/____/________',3)
p('1 Desenvolvimento do exercício',4)
p('Neste exercício, desenvolvi o cálculo do IMC em Java e usei testes com JUnit 5 para comparar os resultados. Também configurei uma pipeline no GitHub Actions e fiz uma alteração proposital na fórmula para demonstrar como os testes identificam um erro.')
p('A classe CalcIMC recebe o peso em quilogramas e a altura em metros. O método calcularIMC(double peso, double altura) retorna o resultado da fórmula:')
p('IMC = peso / (altura * altura)')
p('Com peso de 80 kg e altura de 2 m, o resultado esperado é 20,0. Com 70 kg e 1,75 m, o resultado é aproximadamente 22,8571.')
p('2 Repositório e estrutura do projeto',4)
p('Repositório: https://github.com/recheraphael-Re/aula10devops_IMC')
p('Organizei o código em src/main/java/com/CalcIMC e os testes em src/test/java/com/CalcIMC. O pom.xml contém as configurações do Maven, e o workflow está em .github/workflows/aula10-imc-testes.yml.')
p('Print 1 Repositório com o código e o workflow do CalcIMC',12); box()
p('3 Testes automatizados com JUnit',4,True)
p('Criei a classe CalcIMCTest com dois testes para o cálculo. Usei assertEquals com uma tolerância de 0,0001 para comparar os valores decimais.')
p('O primeiro teste calcula o IMC de 80 kg e 2 m e compara o resultado com 20,0. O segundo usa 70 kg e 1,75 m e compara com 22,857142857.')
p('Para executar somente os testes da calculadora, o comando é:')
p('mvn -Dtest=CalcIMCTest test')
p('Com a fórmula correta, o resultado esperado é: Tests run: 2, Failures: 0, Errors: 0 e BUILD SUCCESS. O comando mvn clean test executa todos os testes existentes no projeto.')
p('Print 2 Execução dos testes do CalcIMC e resultado no terminal ou no Actions',12); box()
p('4 Simulação de erro no cálculo',4,True)
p('Para simular um erro simples, troquei a divisão por multiplicação no retorno do método:')
p('return peso * (altura * altura);')
p('Com essa alteração, o exemplo de 80 kg e 2 m passa a retornar 320,0, enquanto o teste continua esperando 20,0. Assim, o teste consegue apontar que o cálculo está incorreto.')
p('Publiquei essa alteração no commit 4960d56, com a mensagem “Simula erro no calculo de IMC para demonstrar falha no CI”. O resultado esperado dessa versão é uma falha nos testes e BUILD FAILURE.')
p('Resultado registrado no GitHub Actions: ______________________________')
p('Print 3 Pipeline da simulação com o teste que falhou e a comparação dos valores',12); box()
p('5 Correção do cálculo e nova execução',4,True)
p('Depois da simulação, restaurei a fórmula correta:')
p('return peso / (altura * altura);')
p('A correção foi publicada no commit d795857, com a mensagem “Corrige calculo de IMC apos simulacao de falha”. Com essa versão, os resultados voltam a corresponder aos valores esperados pelos testes.')
p('O workflow usa Java 21 e executa mvn -B clean verify. Além de compilar e testar, ele gera o JAR e disponibiliza esse arquivo como artefato quando a execução termina com sucesso. Essa etapa representa a build; a publicação em um servidor não faz parte desse workflow.')
p('Resultado registrado após a correção: _________________________________')
p('Print 4 Pipeline da correção com os testes aprovados e o artefato da build',12); box()
p('6 Respostas às perguntas da aula',4,True)
for h,t in [
('a Comando para limpar o build','O comando é mvn clean. Ele remove a pasta target e os arquivos gerados nas compilações anteriores.'),
('b Comando para executar os testes com Maven','O comando é mvn test. Também posso usar mvn clean test para limpar os arquivos anteriores e executar os testes em seguida.'),
('c Testes unitários antes dos testes E2E','Os testes unitários verificam partes menores do código, são mais rápidos e ajudam a encontrar a origem de um erro. Os testes E2E verificam o fluxo completo e envolvem mais componentes. Por isso, faz sentido corrigir primeiro os problemas nas partes menores.'),
('d Diferença entre Maven e Spring Boot','Maven é uma ferramenta que gerencia dependências, compilação, testes e empacotamento. Spring Boot é um framework que facilita a criação de aplicações Spring, como APIs com servidor embutido. Um projeto Spring Boot também pode usar Maven, como acontece na versão local do CalcIMC.')]:
    p(h,17); p(t)
p('7 Conclusão',4)
p('Com o exercício, entendi como os testes automatizados ajudam a conferir o cálculo e a detectar mudanças incorretas. A simulação deixou clara a diferença entre o valor esperado e o valor retornado. Ao corrigir a fórmula, mantive os mesmos testes para conferir novamente o comportamento do método.')
body.append(sect)
xml=E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(out,'w',ZIP_DEFLATED) as dest:
    for item in z.infolist(): dest.writestr(item,xml if item.filename=='word/document.xml' else z.read(item.filename))
print(out.resolve())
print('Preserved package parts:',len(z.namelist())-1)
