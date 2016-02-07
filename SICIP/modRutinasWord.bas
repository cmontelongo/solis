Attribute VB_Name = "modRutinasWord"
Private Sub CrearContrato()


'https://msdn.microsoft.com/en-us/library/office/aa192487%28v=office.11%29.aspx
'https://support.microsoft.com/es-es/kb/313193
'https://msdn.microsoft.com/en-us/library/6b9478cs.aspx
'https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.selection.aspx
'http://www.xtremevbtalk.com/word-powerpoint-outlook-and-other-office-products/282840-word-automation.html#post1229195
'http://www.xtremevbtalk.com/attachment.php?attachmentid=26680&d=1177831641&
'https://msdn.microsoft.com/en-us/library/8b7k14a4.aspx
'https://norfipc.com/utiles/codigos-ejemplos-macros-para-word-en-visual-basic.php
'http://www.vbforums.com/showthread.php?558056-Word-VBA-Paragraph-formatting
'https://msdn.microsoft.com/en-us/library/office/aa196464%28v=office.11%29.aspx


Dim oword As Object
Dim odoc As Object
Dim oPara1 As Object 'Word.Paragraph
Dim oPara2 As Word.Paragraph
Dim oPara3 As Word.Paragraph
Dim oPara4 As Word.Paragraph


'Set WordAppl = New Word.Application
Set oword = CreateObject("Word.Application")
Set odoc = oword.Documents.Add(DocumentType:=wdNewBlankDocument)

'Insert a paragraph at the beginning of the document.
Set oPara1 = odoc.Content.Paragraphs.Add
oPara1.Range.Text = "CONTRATO DE OBRA A PRECIO ALZADO QUE CELEBRAN POR UNA PARTE METALES TORIZ Y ASOCIADOS, S. A. (A LA QUE EN EL TEXTO SUCESIVO" & _
    "DE ESTE CONTRATO SE HARA REFERENCIA COMO ""EL CLIENTE""), REPRESENTADA POR EL ING. CARLOS ALFONSO TORIZ SOLIS EN SU CARÁCTER DE APODERADO GENERAL Y " & _
    "POR LA OTRA, ING. RICARDO JAVIER SALINAS OVIEDO (QUE SERA REFERIDA EN LO FUTURO EN EL PRESENTE CONTRATO COMO ""EL SUBCONTRATISTA""), REPRESENTADA POR" & _
    "EL ING. RICARDO JAVIER SALINAS OVIEDO EN SU CARÁCTER DE APODERADO GENERAL, DE CONFORMIDAD CON LAS SIGUIENTES DECLARACIONES Y CLAUSULAS."
oPara1.Range.Font.Bold = True
oPara1.Format.SpaceAfter = 12    '24 pt spacing after paragraph.
oPara1.Format.Alignment = 3    'wdAlignParagraphJustify
oPara1.Range.InsertParagraphAfter

oword.Selection.TypeText "D E C L A R A C I O N E S"
oword.Selection.Font.Bold = True
oword.Selection.Format.Alignment = 1 'wdAlignParagraphCenter

oword.Selection.TypeText "Declara ""EL SUBCONTRATISTA"" que es una persona física constituida de conformidad con las leyes aplicables de la República Mexicana, " & _
    "al corriente de sus obligaciones fiscales, laborales y de seguridad social, que está inscrita en el Registro Federal de Contribuyentes, bajo el número " & _
    "SAOR610409G93 con domicilio fiscal ubicado en Ave. Portal No. 225 Col. Portal del Huajuco CP 64989 Monterrey, N.L., contando con " & _
    "la capacidad, medios y elementos propios suficientes para cumplir con sus obligaciones y para llevar a cabo obras industriales de construcción, " & _
    "así como para cumplir con todas aquellas obligaciones que deriven de las relaciones con sus trabajadores, que se encuentra inscrita ante el Instituto " & _
    "Mexicano del Seguro Social con número de Registro Patronal Y403747910-4 Y QUE SE ENCUENTRA DEBIDAMENTE CLASIFICADA CONFORME AL Catálogo de Actividades " & _
    "para la Clasificación de las Empresas en el Seguro de Riesgo de Trabajo, previsto en el Reglamento de la Ley del Seguro Social en Materia de Afiliación, " & _
    "Clasificación de Empresas, Recaudación y Fiscalización."
oword.Selection.Font.Bold = False

'Selection.MoveUp Unit:=wdLine, Count:=2
'    Selection.MoveDown Unit:=wdLine, Count:=1
'    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
'    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
'        .NumberFormat = "%1."
'        .TrailingCharacter = wdTrailingTab
'        .NumberStyle = wdListNumberStyleUppercaseRoman
'        .NumberPosition = CentimetersToPoints(0.63)
'        .Alignment = wdListLevelAlignRight
'        .TextPosition = CentimetersToPoints(1.27)
'        .TabPosition = wdUndefined
'        .ResetOnHigher = 0
'        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .Strikethrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = ""
'    End With
'    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
'    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
'        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
'        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
'        wdWord10ListBehavior
'    Selection.Font.Bold = wdToggle
'    With Selection.ParagraphFormat
'        .LeftIndent = CentimetersToPoints(1.26)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 0
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 6
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpaceMultiple
'        .LineSpacing = LinesToPoints(1.15)
'        .Alignment = wdAlignParagraphLeft
'        .WidowControl = True
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(-0.63)
'        .OutlineLevel = wdOutlineLevelBodyText
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'    End With




odoc.SaveAs FileName:="W:\TestWordDoc.doc"

odoc.Close False
oword.Quit False

Set odoc = Nothing
Set oword = Nothing


    
    

End Sub
Private Sub GeneraContrato()

Dim oword As Word.Application
Dim odoc As Word.Document
Dim oselection As Word.Selection
Dim orange As Word.Range
Dim strTexto As String


Dim objWord As Object
Dim thisDoc As Object
Dim thisRange As Object
Dim varLCount As Variant
Dim thisSelection As Object

Screen.MousePointer = vbHourglass
Set objWord = CreateObject("Word.Application")
Set thisDoc = objWord.Documents.Add
Set thisSelection = objWord.Selection
'Set oselection = oword.Selection

thisDoc.Range.InsertAfter "CONTRATO DE OBRA A PRECIO ALZADO QUE CELEBRAN POR UNA PARTE METALES TORIZ Y ASOCIADOS, S. A. (A LA QUE EN EL TEXTO SUCESIVO" & _
    "DE ESTE CONTRATO SE HARA REFERENCIA COMO ""EL CLIENTE""), REPRESENTADA POR EL ING. CARLOS ALFONSO TORIZ SOLIS EN SU CARÁCTER DE APODERADO GENERAL Y " & _
    "POR LA OTRA, ING. RICARDO JAVIER SALINAS OVIEDO (QUE SERA REFERIDA EN LO FUTURO EN EL PRESENTE CONTRATO COMO ""EL SUBCONTRATISTA""), REPRESENTADA POR" & _
    "EL ING. RICARDO JAVIER SALINAS OVIEDO EN SU CARÁCTER DE APODERADO GENERAL, DE CONFORMIDAD CON LAS SIGUIENTES DECLARACIONES Y CLAUSULAS." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = True

thisDoc.Range.InsertAfter "D E C L A R A C I O N E S" & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 1 'wdAlignParagraphCenter
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 10
thisRange.Font.Bold = True

thisDoc.Range.InsertAfter "Declara ""EL SUBCONTRATISTA"" que es una persona física constituida de conformidad con las leyes aplicables de la República Mexicana, " & _
    "al corriente de sus obligaciones fiscales, laborales y de seguridad social, que está inscrita en el Registro Federal de Contribuyentes, bajo el número " & _
    "SAOR610409G93 con domicilio fiscal ubicado en Ave. Portal No. 225 Col. Portal del Huajuco CP 64989 Monterrey, N.L., contando con " & _
    "la capacidad, medios y elementos propios suficientes para cumplir con sus obligaciones y para llevar a cabo obras industriales de construcción, " & _
    "así como para cumplir con todas aquellas obligaciones que deriven de las relaciones con sus trabajadores, que se encuentra inscrita ante el Instituto " & _
    "Mexicano del Seguro Social con número de Registro Patronal Y403747910-4 Y QUE SE ENCUENTRA DEBIDAMENTE CLASIFICADA CONFORME AL Catálogo de Actividades " & _
    "para la Clasificación de las Empresas en el Seguro de Riesgo de Trabajo, previsto en el Reglamento de la Ley del Seguro Social en Materia de Afiliación, " & _
    "Clasificación de Empresas, Recaudación y Fiscalización." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = False

thisDoc.Range.InsertAfter "Declara ""EL SUBCONTRATISTA"" que con elementos, recursos y personal propio, desea y está en condiciones de llevar a cabo para ""EL CLIENTE""  " & _
    "trabajos especializados de BARANDAL  Y PORTONES, mismos que no son parte de las actividades sustantivas que constituyen el objeto principal de " & _
    """EL CLIENTE"" y que incluyen sin limitar el EQUIPO, MATERIALES, MANO DE OBRA Y CUALESQUIER CONTRATACION necesaria, para la nave identificada como " & _
    "KRAEM y que se ubica en PROLONGACION AV. ALBORADA LOTE 6B, PARQUE INDUSTRIAL FINSA GUADALUPE AEROPUERTO, GUADALUPE N.L. y que en sucesivo " & _
    "denominaremos LA OBRA." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = False

thisDoc.Range.InsertAfter "Declara ""EL  SUBCONTRATISTA"" que ha evaluado los riesgos que tiene o puede llegar a tener el inmueble así como los trabajos a ser llevados " & _
    "a cabo en el mismo, en la ubicación descrita en la declaración II anterior, en términos de la legislación y normatividad aplicable." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = False

thisDoc.Range.InsertAfter "Declara ""EL  SUBCONTRATISTA"" que tiene pleno conocimiento de la Norma Oficial Mexicana ""NOM-031-STPS-2011"", ASI COMO DEL Reglamento  " & _
    "del Seguro Social Obligatorio para los Trabajadores de la Construcción por Obra o Tiempo Determinado, que cuenta con los recursos y " & _
    "procedimientos propios, necesarios y suficientes para su total aplicación en LA OBRA." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = False

thisDoc.Range.InsertAfter "Declara ""EL CLIENTE"" que es una sociedad anónima debidamente constituida de conformidad con las normas aplicables de la Ley General " & _
    "de Sociedades Mercantiles, que está inscrita en el Registro Federal de Contribuyentes, bajo el número MTA820311830, con domicilio fiscal " & _
    "ubicado en Napoleón 3408 col. Estrella, Monterrey N.L. CP 64400, que se encuentra inscrita ante el Instituto Mexicano del Seguro Social " & _
    "bajo el Registro Patronal número D5011350109, que no cuenta ni ha contado con trabajadores contratados por obra o tiempo determinado, " & _
    "y que en su deseo que se lleve a cabo LA OBRA a que se hace referencia en la declaración inmediata precedente." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = False

thisDoc.Range.InsertAfter "En atención a las anteriores declaraciones de las partes, las mismas acuerdan sujetar el presente contrato a las estipulaciones que se contienen en las siguientes." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = False

thisDoc.Range.InsertAfter "C L Á U S U L A S" & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 1 'wdAlignParagraphCenter
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = True

thisDoc.Range.InsertAfter "PRIMERA.- OBJETO Y PRECIO DEL CONTRATO. " & _
    "El subcontratista se obliga en favor del Cliente en llevar a cabo los trabajos de MANO DE OBRA Necesarias para la cimentación de un muro bajo " & _
    "y obra civil para la cimentación de poste y riel (en lo sucesivo referido como la Obra), llevando a cabo tales trabajos a un precio alzado de: Por " & _
    "lo que respecta a Mano de Obra y Herramientas: por un monto de 455,980.00 pesos  (Cuatrocientos cincuenta y cinco mil novecientos ochenta pesos 00/100 M.N.) " & _
    "más IVA." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = False

thisDoc.Range.InsertAfter "A dichas cantidades (conjuntamente en lo sucesivo referidas como el ""Precio"") deberá agregarse el importe del Impuesto al Valor Agregado correspondiente." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = False

'oword
'odoc.Font

'odoc.Range.sele
strTexto = "SEGUNDA.- RESPONSABILIDADES DE EL SUBCONTRATISTA. "

'oselection.Font
'thisSelection.Font
'oselection.MoveUp
   ' Selection.MoveUp Unit:=wdLine, Count:=7
   ' Selection.MoveRight Unit:=wdCharacter, Count:=18, Extend:=wdExtend
   ' Selection.Font.Bold = wdToggle
'oselection.MoveEnd

thisDoc.Range.InsertAfter strTexto
thisSelection.MoveUp Unit:=wdLine, Count:=1
thisSelection.MoveRight Unit:=wdCharacter, Count:=Len(strTexto), Extend:=wdExtend
thisSelection.Font.Bold = wdToggle
thisDoc.Range.InsertAfter "El Subcontratista se obliga a llevar a cabo los trabajos en que se hace consistir la obra, precisamente con equipo, herramienta, " & _
    "materiales y trabajadores propios, por lo que desde ahora se obliga el Subcontratista a sacar al Cliente en paz y a salvo de cualquier " & _
    "reclamación que cualquier empresa, trabajador o sindicato contratado por el Subcontratista en relación con la Obra, llegare a hacerle, " & _
    "así como a resarcir al Cliente de cualquier erogación que éste tuviere que hacer con motivo de tales reclamaciones, incluyendo todo gasto " & _
    "y honorario que el Cliente llegare a erogar con motivo de su defensa." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8


thisDoc.Range.InsertAfter "El Subcontratista reconoce que no existe ni existirá relación de trabajo entre el Cliente y el personal que intervenga en el desarrollo " & _
    "de los trabajos en que se hace consistir la Obra y por tanto el Subcontratista asume toda responsabilidad frente a dicho personal, los sindicatos " & _
    "con quien contrate y a los que pertenezcan los trabajadores, las autoridades laborales en general, El Instituto del Fondo Nacional de la Vivienda " & _
    "de los Trabajadores (en lo sucesivo referido como el ""INFONAVIT"") y el Instituto Mexicano del Seguro Social (En lo sucesivo referido como el ""IMSS""), " & _
    "quedando obligada el Subcontratista a sacar en paz y a salvo al Cliente en caso de cualquier reclamación, demanda o procedimiento que se entable " & _
    "en contra del Cliente por los trabajadores que intervengan en el desarrollo de los trabajos, sus sindicatos, las autoridades laborales en general, " & _
    "el INFONAVIT o el IMSS sin importar de que índole se trate o si tales reclamaciones, demandas o procedimientos se declaran o no fundados y " & _
    "procedentes por las autoridades competentes y en consecuencia queda obligada el Subcontratista a indemnizar y resarcir al Cliente de cualquier erogación " & _
    "que en su caso éste tuviere que hacer con motivo de tales reclamaciones, demandas o procedimientos, incluyendo todo y gasto y honorarios que el " & _
    "cliente llegare a hacer con motivo de su defensa, en términos de la Cláusula Decima Octava del presente contrato. Es responsabilidad exclusiva " & _
    "de el Subcontratista y será a su cargo y costa la inscripción en el IMSS mediante el formato SATIC-02, de los trabajadores que en su caso tenga " & _
    "o llegase a contratar para efectivos de la Obra así como el cumplimiento de cualesquiera obligaciones en materia laboral, de seguridad e higiene, " & _
    "retiro, vivienda de los trabajadores y demás que deriven de la legislación aplicable con respecto a sus trabajadores; el cliente tendrá derecho " & _
    "de verificar razonablemente que el Subcontratista éste en cumplimiento con lo anterior." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = False
thisDoc.Range.InsertAfter "Es responsabilidad exclusiva de el Subcontratista y será a su cargo y costa, la adquisición y uso de todos y cada uno de los elementos de " & _
    "equipo de seguridad que deban utilizar sus trabajadores que intervenga en el desarrollo de los trabajos en que se hace consistir la Obra; " & _
    "obligándose a este respecto, el Subcontratista a sujetarse a las leyes, reglamentos y normas legales de seguridad, higiene y medio ambiente " & _
    "incluyendo el Reglamento Federal de Seguridad, Higiene y Medio Ambiente de Trabajo, el Reglamento del Seguro Social obligatorio para los " & _
    "trabajadores de la Construcción por Obra o tiempo Determinado y cualesquiera otras que sean legalmente aplicables, y que en su caso determinen " & _
    "las autoridades laborales y administrativas competentes. Igualmente será responsabilidad exclusiva del Subcontratista la guarda y custodia " & _
    "de todo equipo de seguridad así como de toda herramienta que se utilice por el personal que intervenga en los trabajos de qué trata este contrato. " & _
    "En adicción a lo anterior, el Cliente tendrá derecho en todo momento de dar por terminado el presente contrato sin necesidad de declaración " & _
    "judicial y sin relevar al Subcontratista de responder, indemnizar y sacar en paz y a salvo al Cliente en términos del presente, por cualesquiera " & _
    "violaciones a la legislación en materia de seguridad, higiene y medio ambiente." & vbCrLf
varLCount = thisDoc.Paragraphs.Count - 1
Set thisRange = thisDoc.Paragraphs(varLCount).Range
thisRange.ParagraphFormat.Alignment = 3    'wdAlignParagraphJustify
thisRange.Font.Name = "Arial"
thisRange.Font.Size = 8
thisRange.Font.Bold = False



thisDoc.SaveAs ("W:\Repo.doc")

thisDoc.Close False
objWord.Quit False

Set thisDoc = Nothing
Set objWord = Nothing

Screen.MousePointer = vbDefault
End Sub

Private Sub ImprimeContrato()

Dim objWord As Object
Dim thisDoc As Object
Dim thisSelection As Object

Screen.MousePointer = vbHourglass
Set objWord = CreateObject("Word.Application")
Set thisDoc = objWord.Documents.Add
Set thisSelection = objWord.Selection


    thisSelection.Font.Bold = wdToggle
    thisSelection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    'thisSelection.Paste
    thisSelection.TypeParagraph
    thisSelection.Font.Bold = wdToggle
    thisSelection.Font.Size = 8
    thisSelection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    thisSelection.TypeText Text:="declaraciones"
    thisSelection.TypeParagraph
    thisSelection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleUppercaseRoman
        .NumberPosition = CentimetersToPoints(0.63)
        .Alignment = wdListLevelAlignRight
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .Strikethrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    thisSelection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
    thisSelection.TypeText Text:="declara"
    thisSelection.TypeParagraph
    thisSelection.TypeText Text:="que es"
    thisSelection.TypeParagraph
    thisSelection.TypeText Text:="una persona"
    thisSelection.TypeParagraph
    thisSelection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    thisSelection.TypeText Text:="por lo que debe ser asi"
    thisSelection.TypeParagraph
    thisSelection.Font.Bold = wdToggle
    thisSelection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    thisSelection.TypeText Text:="anuncions"
    thisSelection.TypeParagraph
    thisSelection.Font.Bold = wdToggle
    thisSelection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    thisSelection.Font.Bold = wdToggle
    thisSelection.TypeText Text:="clausula primera"
    thisSelection.Font.Bold = wdToggle
    thisSelection.TypeText Text:=" que no debe de ir por aqui"
    thisSelection.TypeParagraph
    thisSelection.TypeText Text:="no se debe hacer"


thisDoc.SaveAs ("W:\Repo.doc")

thisDoc.Close False
objWord.Quit False

Set thisDoc = Nothing
Set objWord = Nothing

Screen.MousePointer = vbDefault

End Sub
Obra



