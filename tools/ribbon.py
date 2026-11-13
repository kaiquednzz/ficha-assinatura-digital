"""Injeta a guia 'Assinaturas' (customUI) no Coleta.xlam. Uso: python ribbon.py <entrada.xlam> <saida.xlam>"""
import sys, zipfile

UI = """<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
 <ribbon><tabs><tab id="tabAssin" label="Assinaturas">
  <group id="grpFicha" label="Ficha da faca">
   <button id="btnIniciar" label="Iniciar coleta" size="large" imageMso="SignatureLineInsert" onAction="BtnIniciar"
           screentip="Exporta a ficha da faca em D2 e abre o link para o celular"/>
   <button id="btnFinalizar" label="Finalizar coleta" size="large" imageMso="FileSaveAsPdfOrXps" onAction="BtnFinalizar"
           screentip="Aplica as assinaturas e gera o PDF assinado"/>
  </group>
 </tab></tabs></ribbon>
</customUI>"""
REL = ('<Relationship Id="rIdColetaUI" '
       'Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" '
       'Target="customUI/customUI.xml"/>')

src, dst = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "_rels/.rels":
            assert b"rIdColetaUI" not in data
            data = data.replace(b"</Relationships>", REL.encode() + b"</Relationships>")
        zout.writestr(item, data)
    zout.writestr("customUI/customUI.xml", UI)
print("customUI injetado")
