#include "Totvs.ch"

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

Class YExcel From custom.excel.Xlsx
	Method New()
EndClass

Method New(cNomeFile, cFileOpen, cTipo) Class YExcel
	_Super:New(cNomeFile, cFileOpen, cTipo)
Return self

Class YExcel_Style From custom.excel.Style
	Method New()
EndClass

Method New(oPai, oExcel) Class YExcel_Style
	_Super:New(oPai, oExcel)
Return self

Class YExcel_StyleRules From custom.excel.StyleRules
	Method New()
EndClass

Method New(oExcel) Class YExcel_StyleRules
	_Super:New(oExcel)
Return self

Class YExcel_RegraLinha From custom.excel.RegraLinha
	Method New()
EndClass

Method New(bBloco, aRegra, oExcel) Class YExcel_RegraLinha
	_Super:New(bBloco, aRegra, oExcel)
Return self

Class YExcelTag From custom.excel.Tag
	Method New()
EndClass

Method New(cNome, xValor, oAtributo, oExcel) Class YExcelTag
	_Super:New(cNome, xValor, oAtributo, oExcel)
Return self

Class YExcel_Table From custom.excel.Table
	Method New()
EndClass

Method New(oyExcel, nLinha, nColuna, cNome) Class YExcel_Table
	_Super:New(oyExcel, nLinha, nColuna, cNome)
Return self

Class YExcel_DateTime From custom.excel.DateTime
	Method New()
EndClass

Method New(dData, cTime, nData, nDec8, cDataUTC) Class YExcel_DateTime
	_Super:New(dData, cTime, nData, nDec8, cDataUTC)
Return self

Class YExcelVar From custom.excel.Var
	Method New()
EndClass

Method New(cTipo) Class YExcelVar
	_Super:Var():New(cTipo)
Return self

Class YExcelfunction From custom.excel.Formula
	Method New()
EndClass

Method New() Class YExcelfunction
	_Super:Formula():New()
Return self
