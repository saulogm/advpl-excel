#include "Totvs.ch"
#include "Fileio.ch"
#Include "ParmType.ch"
#INCLUDE "DBSTRUCT.CH"
#INCLUDE "DBINFO.CH"

Static cAr7Zip	:= GetSrvProfString("7zip","")
Static cRootPath
Static c7Zip	:= "Z"

/*/{Protheus.doc} YExcel
Gera planilha excel
@author Saulo Gomes Martins
@since 03/05/2017
@version p11
@param cNomeFile, characters, Nome do arquivo sem extensao
@param cFileOpen, characters, (Opcional) Arquivo para abertura e edicao
@param cTipo, characters, (Opcional) Tipo de criacao (BD=banco de dados padrao;TMPDB=banco temporario;MEM=memoria)
@type class
/*/

User Function YExcel()
Return .T.

Class YExcel
	Method New()
	Method ClassName()
EndClass

Method New(cNomeFile, cFileOpen, cTipo) Class YExcel
	Local oObj := custom.tools.excel.Xlsx():New(cNomeFile, cFileOpen, cTipo)
Return oObj

Method ClassName() Class YExcel
Return "YEXCEL"

Class YExcel_Style
	Method New()
	Method ClassName()
EndClass

Method New(oPai, oExcel) Class YExcel_Style
	Local oObj := custom.tools.excel.Style():New(oPai, oExcel)
Return oObj

Method ClassName() Class YExcel_Style
Return "YEXCEL_STYLE"

Class YExcel_StyleRules
	Method New()
	Method ClassName()
EndClass

Method New(oExcel) Class YExcel_StyleRules
	Local oObj := custom.tools.excel.StyleRules():New(oExcel)
Return oObj

Method ClassName() Class YExcel_StyleRules
Return "YEXCEL_STYLERULES"

Class YExcel_RegraLinha
	Method New()
	Method ClassName()
EndClass

Method New(bBloco, aRegra, oExcel) Class YExcel_RegraLinha
	Local oObj := custom.tools.excel.RegraLinha():New(bBloco, aRegra, oExcel)
Return oObj

Method ClassName() Class YExcel_RegraLinha
Return "YEXCEL_REGRALINHA"

Class YExcelTag
	Method New()
	Method ClassName()
EndClass

Method New(cNome, xValor, oAtributo, oExcel) Class YExcelTag
	Local oObj := custom.tools.excel.Tag():New(cNome, xValor, oAtributo, oExcel)
Return oObj

Method ClassName() Class YExcelTag
Return "YEXCELTAG"

Class YExcel_Table
	Method New()
	Method ClassName()
EndClass

Method New(oyExcel, nLinha, nColuna, cNome) Class YExcel_Table
	Local oObj := custom.tools.excel.Table():New(oyExcel, nLinha, nColuna, cNome)
Return oObj

Method ClassName() Class YExcel_Table
Return "YEXCEL_TABLE"

Class YExcel_DateTime
	Method New()
	Method ClassName()
EndClass

Method New(dData, cTime, nData, nDec8, cDataUTC) Class YExcel_DateTime
	Local oObj := custom.tools.excel.DateTime():New(dData, cTime, nData, nDec8, cDataUTC)
Return oObj

Method ClassName() Class YExcel_DateTime
Return "YEXCEL_DATETIME"

Class YExcelVar
	Method New()
	Method ClassName()
EndClass

Method New(cTipo) Class YExcelVar
	Local oObj := custom.tools.excel.Var():New(cTipo)
Return oObj

Method ClassName() Class YExcelVar
Return "YEXCELVAR"

Class YExcelfunction
	Method New()
	Method ClassName()
EndClass

Method New() Class YExcelfunction
	Local oObj := custom.tools.excel.Formula():New()
Return oObj

Method ClassName() Class YExcelfunction
Return "YEXCELFUNCTION"