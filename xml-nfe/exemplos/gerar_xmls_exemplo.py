"""Gera dois XMLs de NF-e fictícios para experimentar a ferramenta.

    python gerar_xmls_exemplo.py            -> grava ao lado deste script
    python gerar_xmls_exemplo.py <pasta>    -> grava na pasta indicada

Os arquivos seguem a estrutura mínima nfeProc/NFe/infNFe com o namespace oficial
(http://www.portalfiscal.inf.br/nfe), itens com CFOP e um bloco Signature de mentira,
para dar para conferir que a troca de CFOP preserva o namespace e a assinatura.
Não há CNPJ, chave de acesso nem qualquer dado de empresa real: os nomes são inventados.

O nome dos arquivos começa com o número da nota e um traço baixo ("1001_..."), que é
o padrão que a aba "Separar por número" procura.
"""
import os
import sys
from pathlib import Path

PASTA = Path(os.path.abspath(__file__)).parent

NOTAS = [
    {
        "numero": "1001", "serie": "1", "emissao": "2026-07-15T10:12:00-03:00",
        "emitente": "Empresa Exemplo Ltda", "destinatario": "Comércio Modelo Ltda",
        "itens": [
            ("P0001", "Parafuso sextavado M6 x 40 zincado", "73181500", "5102", "UN", "1200.0000", "0.4200", "504.00"),
            ("P0002", "Porca sextavada M6 zincada", "73181600", "5102", "UN", "1500.0000", "0.1800", "270.00"),
            ("P0009", "Disco de corte 7 pol", "68042290", "5405", "UN", "60.0000", "9.8000", "588.00"),
        ],
    },
    {
        "numero": "1002", "serie": "1", "emissao": "2026-07-16T14:40:00-03:00",
        "emitente": "Empresa Exemplo Ltda", "destinatario": "Indústria Fictícia S.A.",
        "itens": [
            ("P0004", "Chapa de aço carbono 1,5 mm 1000 x 2000", "72085200", "5102", "PC", "18.0000", "312.5000", "5625.00"),
            ("P0001", "Parafuso sextavado M6 x 40 zincado", "73181500", "5102", "UN", "500.0000", "0.4200", "210.00"),
        ],
    },
]


def xml_da_nota(nota: dict) -> str:
    ident = f"NFe-EXEMPLO-{nota['serie']}-{nota['numero']}"
    itens = []
    total = 0.0
    for n, (cprod, xprod, ncm, cfop, un, q, vun, vprod) in enumerate(nota["itens"], start=1):
        total += float(vprod)
        itens.append(f"""      <det nItem="{n}">
        <prod>
          <cProd>{cprod}</cProd>
          <xProd>{xprod}</xProd>
          <NCM>{ncm}</NCM>
          <CFOP>{cfop}</CFOP>
          <uCom>{un}</uCom>
          <qCom>{q}</qCom>
          <vUnCom>{vun}</vUnCom>
          <vProd>{vprod}</vProd>
        </prod>
        <imposto>
          <ICMS><ICMS00><orig>0</orig><CST>00</CST><pICMS>18.00</pICMS></ICMS00></ICMS>
        </imposto>
      </det>""")
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">
  <NFe>
    <infNFe Id="{ident}" versao="4.00">
      <ide>
        <cUF>35</cUF>
        <natOp>VENDA DE MERCADORIA</natOp>
        <mod>55</mod>
        <serie>{nota['serie']}</serie>
        <nNF>{nota['numero']}</nNF>
        <dhEmi>{nota['emissao']}</dhEmi>
        <tpNF>1</tpNF>
        <tpAmb>2</tpAmb>
      </ide>
      <emit>
        <xNome>{nota['emitente']}</xNome>
        <enderEmit><xMun>Cidade Exemplo</xMun><UF>SP</UF></enderEmit>
      </emit>
      <dest>
        <xNome>{nota['destinatario']}</xNome>
        <enderDest><xMun>Vila Modelo</xMun><UF>SP</UF></enderDest>
      </dest>
{chr(10).join(itens)}
      <total>
        <ICMSTot>
          <vProd>{total:.2f}</vProd>
          <vNF>{total:.2f}</vNF>
        </ICMSTot>
      </total>
    </infNFe>
    <Signature xmlns="http://www.w3.org/2000/09/xmldsig#">
      <SignedInfo>
        <CanonicalizationMethod Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>
        <SignatureMethod Algorithm="http://www.w3.org/2000/09/xmldsig#rsa-sha1"/>
        <Reference URI="#{ident}">
          <DigestMethod Algorithm="http://www.w3.org/2000/09/xmldsig#sha1"/>
          <DigestValue>RESUMO-DE-EXEMPLO-NAO-E-REAL=</DigestValue>
        </Reference>
      </SignedInfo>
      <SignatureValue>ASSINATURA-DE-EXEMPLO-NAO-E-REAL=</SignatureValue>
      <KeyInfo>
        <X509Data>
          <X509Certificate>CERTIFICADO-DE-EXEMPLO-NAO-E-REAL=</X509Certificate>
        </X509Data>
      </KeyInfo>
    </Signature>
  </NFe>
  <protNFe versao="4.00">
    <infProt>
      <tpAmb>2</tpAmb>
      <cStat>100</cStat>
      <xMotivo>Autorizado o uso da NF-e (ambiente de exemplo)</xMotivo>
    </infProt>
  </protNFe>
</nfeProc>
"""


def gerar(pasta: Path) -> list:
    pasta.mkdir(parents=True, exist_ok=True)
    gravados = []
    for nota in NOTAS:
        destino = pasta / f"{nota['numero']}_nfe-exemplo.xml"
        destino.write_text(xml_da_nota(nota), encoding="utf-8", newline="\n")
        gravados.append(destino)
    return gravados


if __name__ == "__main__":
    alvo = Path(sys.argv[1]) if len(sys.argv) > 1 else PASTA
    for arquivo in gerar(alvo):
        print("Gravado:", arquivo)
