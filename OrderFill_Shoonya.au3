#include "CUIAutomation2.au3"
#include <FileConstants.au3>

;-------------------- Create Variables and Read from INI --------------

$sFilePath = @ScriptDir & "\config.ini"

Local $Exchange = IniRead($sFilePath, "Value", "Exchange", "Default Value")
Local $Inst = IniRead($sFilePath, "Value", "Inst", "Default Value")
Local $Symbol = IniRead($sFilePath, "Value", "Symbol", "Default Value")
Local $Strike = IniRead($sFilePath, "Value", "Strike", "Default Value")
Local $Ordertype = IniRead($sFilePath, "Value", "OrderType", "Default Value")
Local $Qty = IniRead($sFilePath, "Value", "Qty", "Default Value")
Local $BuyPrice = IniRead($sFilePath, "Value", "BuyPrice", "Default Value")
Local $SellPrice = IniRead($sFilePath, "Value", "SellPrice", "Default Value")
Local $SL_Trigger = IniRead($sFilePath, "Value", "SL Trigger", "Default Value")
Local $Product = IniRead($sFilePath, "Value", "Product", "Default Value")
Local $CO_Trigger = IniRead($sFilePath, "Value", "CO Trigger", "Default Value")
Local $BO_TP = IniRead($sFilePath, "Value", "BO TakeProfit", "Default Value")
Local $BO_stop = IniRead($sFilePath, "Value", "BO StopLoss", "Default Value")
Local $BO_Trail = IniRead($sFilePath, "Value", "BO Trailing_ONOFF", "Default Value")
Local $BO_TrailVal = IniRead($sFilePath, "Value", "BO Trailing", "Default Value")
Local $Order_Confirm = IniRead($sFilePath, "Value", "Order Confirm", "Default Value")
Local $BuySell = IniRead($sFilePath, "Value", "Buy Sell", "Default Value")

WinActivate("SHOONYA")
WinActivate("shoonya")

if $BuySell = "BUY"	Then
Send("{F1}")
WinActivate("Buy Order")
ElseIf $BuySell = "SELL" Then
Send("{F2}")
WinActivate("Sell Order")
EndIf

Sleep(500)


;Create UI Automation object
Local $oUIAutomation = ObjCreateInterface( $sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtagIUIAutomation )
;Get Desktop element
Local $pDesktop, $oDesktop
$oUIAutomation.GetRootElement( $pDesktop )
$oDesktop = ObjCreateInterface( $pDesktop, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

;Find Shoonya
Local $Shoonya, $pWin_RRMdi, $oWin_RRMdi
$oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "RRMdi", $Shoonya )
$oDesktop.FindFirst( $TreeScope_Children, $Shoonya, $pWin_RRMdi )
$oWin_RRMdi = ObjCreateInterface( $pWin_RRMdi, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

;Find Buy Window
Local $WinBuyOrder,$pBuyOrderView, $oBuyOrderView
$oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "WinBuyOrderView", $WinBuyOrder )
$oWin_RRMdi.FindFirst( $TreeScope_Descendants, $WinBuyOrder, $pBuyOrderView )
$oBuyOrderView = ObjCreateInterface( $pWin_RRMdi, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

;Find Custom BuyWindows (insdie element)
Local $OrderWindow, $pUcNewBuySell, $oUcNewBuySell
$oUIAutomation.CreatePropertyCondition( $UIA_ClassNamePropertyId, "UcNewBuySell", $OrderWindow )
$oBuyOrderView.FindFirst( $TreeScope_Descendants, $OrderWindow, $pUcNewBuySell )
$oUcNewBuySell = ObjCreateInterface( $pUcNewBuySell, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

;------------------- automation ------Exchange -----------------------------------------------------------

Local $pExchange
$oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "CMBExchange", $pExchange )

Local $pCombo_Ex, $oCombo_Ex
$oBuyOrderView.FindFirst( $TreeScope_Descendants, $pExchange, $pCombo_Ex )
$oCombo_Ex = ObjCreateInterface( $pCombo_Ex, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

Local $pExpandCollapsePattern1, $oExpandCollapsePattern1
$oCombo_Ex.GetCurrentPattern( $UIA_ExpandCollapsePatternId, $pExpandCollapsePattern1 )
$oExpandCollapsePattern1 = ObjCreateInterface( $pExpandCollapsePattern1, $sIID_IUIAutomationExpandCollapsePattern, $dtagIUIAutomationExpandCollapsePattern )
$oExpandCollapsePattern1.Expand()

$oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, $Exchange, $pExchange )

Local $pList_ex, $oList_ex
$oCombo_Ex.FindFirst( $TreeScope_Descendants, $pExchange, $pList_ex )
$oList_ex = ObjCreateInterface( $pList_ex, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

Local $pSelection_ex, $oSelection_ex
$oList_ex.GetCurrentPattern( $UIA_SelectionItemPatternId, $pSelection_ex )
$oSelection_ex = ObjCreateInterface( $pSelection_ex, $sIID_IUIAutomationSelectionItemPattern, $dtagIUIAutomationSelectionItemPattern )
$oSelection_ex.Select()

;-------------------- automation ------Instrument -----------------------------------------------------------

Local $CMBInstrument
$oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "CMBInstrument", $CMBInstrument )
Local $pComboBox_inst, $oComboBox_inst
$oUcNewBuySell.FindFirst( $TreeScope_Descendants, $CMBInstrument, $pComboBox_inst )
$oComboBox_inst = ObjCreateInterface( $pComboBox_inst, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

Local $pExpandCollapsePattern2, $oExpandCollapsePattern2
$oComboBox_inst.GetCurrentPattern( $UIA_ExpandCollapsePatternId, $pExpandCollapsePattern2 )
$oExpandCollapsePattern2 = ObjCreateInterface( $pExpandCollapsePattern2, $sIID_IUIAutomationExpandCollapsePattern, $dtagIUIAutomationExpandCollapsePattern )
$oExpandCollapsePattern2.Expand()

Local $pInst_type
$oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, $Inst, $pInst_type )

Local $pInst_list, $oInst_list
$oComboBox_inst.FindFirst( $TreeScope_Descendants, $pInst_type, $pInst_list )
$oInst_list = ObjCreateInterface( $pInst_list, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

Local $pInst_Selection, $oInst_Selection
$oInst_list.GetCurrentPattern( $UIA_SelectionItemPatternId, $pInst_Selection )
$oInst_Selection = ObjCreateInterface( $pInst_Selection, $sIID_IUIAutomationSelectionItemPattern, $dtagIUIAutomationSelectionItemPattern )
$oInst_Selection.Select()

;-------------------- automation ------ Symbol --------------

Local $pSymbol
$oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "CMBSecurities", $pSymbol )

Local $pSymbol_edit, $oSymbol_edit
$oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pSymbol, $pSymbol_edit )
$oSymbol_edit = ObjCreateInterface( $pSymbol_edit, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )
$oSymbol_edit.Setfocus()
Send($symbol)

;-------------------- automation ------ Strike Price --------------

 Local $pOptStrike
 $oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "CMBStrike", $pOptStrike )

 Local $pCombo_Strike, $oCombo_Strike
 $oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pOptStrike, $pCombo_Strike )
 $oCombo_Strike = ObjCreateInterface( $pCombo_Strike, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

 Local $pValue_Strike, $oValue_Strike
 $oCombo_Strike.GetCurrentPattern( $UIA_ValuePatternId, $pValue_Strike )
 $oValue_Strike = ObjCreateInterface( $pValue_Strike, $sIID_IUIAutomationValuePattern, $dtagIUIAutomationValuePattern )
 $oValue_Strike.SetValue($Strike)

;-------------------- automation ------ Order Type --------------

 Local $CMBorder
 $oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "CMBorder", $CMBorder )
 Local $pComboBox_Order, $oComboBox_Order, $pOrder_Type
 $oUcNewBuySell.FindFirst( $TreeScope_Descendants, $CMBorder, $pComboBox_Order )
 $oComboBox_Order = ObjCreateInterface( $pComboBox_Order, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

 Local $pExpandCollapsePattern3, $oExpandCollapsePattern3
 $oComboBox_Order.GetCurrentPattern( $UIA_ExpandCollapsePatternId, $pExpandCollapsePattern3 )
 $oExpandCollapsePattern3 = ObjCreateInterface( $pExpandCollapsePattern3, $sIID_IUIAutomationExpandCollapsePattern, $dtagIUIAutomationExpandCollapsePattern )
 $oExpandCollapsePattern3.Expand()

 $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, $ordertype, $pOrder_Type )

 Local $pOrder_list, $oOrder_list
 $oComboBox_Order.FindFirst( $TreeScope_Descendants, $pOrder_Type, $pOrder_list )
 $oOrder_list = ObjCreateInterface( $pOrder_list, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

 Local $pOrder_Selection, $oOrder_Selection
 $oOrder_list.GetCurrentPattern( $UIA_SelectionItemPatternId, $pOrder_Selection )
 $oOrder_Selection = ObjCreateInterface( $pOrder_Selection, $sIID_IUIAutomationSelectionItemPattern, $dtagIUIAutomationSelectionItemPattern )
 $oOrder_Selection.Select()

 ;------------------- automation ------ Quantity --------------

 Local $pQty
 $oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "lblQuantity", $pQty )

 Local $pTxtQty, $oTxtQty
 $oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pQty, $pTxtQty )
 $oTxtQty = ObjCreateInterface( $pTxtQty, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

 Local $pRawViewWalker_Qty, $oRawViewWalker_Qty
 $oUIAutomation.RawViewWalker( $pRawViewWalker_Qty )
 $oRawViewWalker_Qty = ObjCreateInterface( $pRawViewWalker_Qty, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)

 Local $pEdit_Qty, $oEdit_Qty
 $oRawViewWalker_Qty.GetNextSiblingElement( $oTxtQty, $pEdit_Qty )
 $oEdit_Qty = ObjCreateInterface( $pEdit_Qty, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

 Local $pValue_Qty, $oValue_Qty
 $oEdit_Qty.GetCurrentPattern( $UIA_ValuePatternId, $pValue_Qty )
 $oValue_Qty = ObjCreateInterface( $pValue_Qty, $sIID_IUIAutomationValuePattern, $dtagIUIAutomationValuePattern )
 $oValue_Qty.SetValue($qty)


 ;------------------- automation ------ Buy Price --------------
If $Ordertype = "[L, Limit]" Or "[SL, SL]" Then
 Local $pPrice
 $oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "lblPrice", $pPrice )

 Local $pTxtPrice, $oTxtPrice
 $oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pPrice, $pTxtPrice )
 $oTxtPrice = ObjCreateInterface( $pTxtPrice, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

 Local $pRawViewWalker_price, $oRawViewWalker_price
 $oUIAutomation.RawViewWalker( $pRawViewWalker_price )
 $oRawViewWalker_price = ObjCreateInterface( $pRawViewWalker_price, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)

 Local $pEdit_price, $oEdit_price
 $oRawViewWalker_price.GetNextSiblingElement( $oTxtPrice, $pEdit_price )
 $oEdit_price = ObjCreateInterface( $pEdit_price, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

 Local $pValue_Price, $oValue_Price
 $oEdit_price.GetCurrentPattern( $UIA_ValuePatternId, $pValue_Price )
 $oValue_Price = ObjCreateInterface( $pValue_Price, $sIID_IUIAutomationValuePattern, $dtagIUIAutomationValuePattern )


   if $BuySell = "BUY"	Then
	  $oValue_Price.SetValue($BuyPrice)
	  ElseIf $BuySell = "SELL" Then
	  $oValue_Price.SetValue($SellPrice)
   EndIf


EndIf
 ;------------------- automation ------ SL trigger Price --------------
If $Ordertype = "[SL, SL]" Or "[SL-M, SL-M]" Then

   Local $pSL_trig
   $oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "lblTriggerPrice", $pSL_trig )

   Local $pTxtSL_trig, $oTxtSL_trig
   $oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pSL_trig, $pTxtSL_trig )
   $oTxtSL_trig = ObjCreateInterface( $pTxtSL_trig, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

   Local $pRawViewWalker_slTrig, $oRawViewWalker_slTrig
   $oUIAutomation.RawViewWalker( $pRawViewWalker_slTrig )
   $oRawViewWalker_slTrig = ObjCreateInterface( $pRawViewWalker_slTrig, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)

   Local $pEdit_SL_Trig, $oEdit_SL_Trig
   $oRawViewWalker_slTrig.GetNextSiblingElement( $oTxtSL_trig, $pEdit_SL_Trig )
   $oEdit_SL_Trig = ObjCreateInterface( $pEdit_SL_Trig, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

   Local $pValue_SL_trig, $oValue_SL_trig
   $oEdit_SL_Trig.GetCurrentPattern( $UIA_ValuePatternId, $pValue_SL_trig )
   $oValue_SL_trig = ObjCreateInterface( $pValue_SL_trig, $sIID_IUIAutomationValuePattern, $dtagIUIAutomationValuePattern )
   $oValue_SL_trig.SetValue($SL_Trigger)

EndIf

;------------------- automation ------ Product --------------
Local $pProduct
$oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "CMBProduct", $pProduct )

Local $pCombo_Product, $oCombo_Product
$oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pProduct, $pCombo_Product )
$oCombo_Product = ObjCreateInterface( $pCombo_Product, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

Local $pExpandCollapsePattern4, $oExpandCollapsePattern4
$oCombo_Product.GetCurrentPattern( $UIA_ExpandCollapsePatternId, $pExpandCollapsePattern4 )
$oExpandCollapsePattern4 = ObjCreateInterface( $pExpandCollapsePattern4, $sIID_IUIAutomationExpandCollapsePattern, $dtagIUIAutomationExpandCollapsePattern )
$oExpandCollapsePattern4.Expand()

$oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, $Product, $pProduct )

Local $pList_Product, $oList_Product
$oCombo_Product.FindFirst( $TreeScope_Descendants, $pProduct, $pList_Product )
$oList_Product = ObjCreateInterface( $pList_Product, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

Local $pSelection_Product, $oSelection_Product
$oList_Product.GetCurrentPattern( $UIA_SelectionItemPatternId, $pSelection_Product )
$oSelection_Product = ObjCreateInterface( $pSelection_Product, $sIID_IUIAutomationSelectionItemPattern, $dtagIUIAutomationSelectionItemPattern )

$oSelection_Product.Select()

 ;------------------- automation ------ Cover Order Trigger Using IF --------------

If $Product = "[CO, CO]" Then
   Local $pCO_trig
   $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "     Trigger Price: ", $pCO_trig )

   Local $pText_coTrig, $oText_coTrig
   $oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pCO_trig, $pText_coTrig )
   $oText_coTrig = ObjCreateInterface( $pText_coTrig, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

   Local $pRawViewWalker_COTrig, $oRawViewWalker_COTrig
   $oUIAutomation.RawViewWalker( $pRawViewWalker_COTrig )
   $oRawViewWalker_COTrig = ObjCreateInterface( $pRawViewWalker_COTrig, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)

   Local $pEdit_CO_Trig, $oEdit_CO_Trig
   $oRawViewWalker_COTrig.GetNextSiblingElement( $oText_coTrig, $pEdit_CO_Trig )
   $oEdit_CO_Trig = ObjCreateInterface( $pEdit_CO_Trig, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

   Local $pValue_CO_trig, $oValue_CO_trig
   $oEdit_CO_Trig.GetCurrentPattern( $UIA_ValuePatternId, $pValue_CO_trig )
   $oValue_CO_trig = ObjCreateInterface( $pValue_CO_trig, $sIID_IUIAutomationValuePattern, $dtagIUIAutomationValuePattern )
   $oValue_CO_trig.SetValue($CO_Trigger)

EndIf

 ;------------------- automation ------ BO --------------

If $Product = "[BO, BO]" Then


   ;------------------- automation ------ BO  SquareOff/ Take Profit --------------

   Local $pBO_Sq
   $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "SqOff:  ", $pBO_Sq )

   Local $pText_BOsq, $oText_BOsq
   $oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pBO_Sq, $pText_BOsq )
   $oText_BOsq = ObjCreateInterface( $pText_BOsq, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

   Local $pRawViewWalker_BOSq, $oRawViewWalker_BOSq
   $oUIAutomation.RawViewWalker( $pRawViewWalker_BOSq )
   $oRawViewWalker_BOSq = ObjCreateInterface( $pRawViewWalker_BOSq, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)

   Local $pEdit_BO_sq, $oEdit_BO_sq
   $oRawViewWalker_BOSq.GetNextSiblingElement( $oText_BOsq, $pEdit_BO_sq )
   $oEdit_BO_sq = ObjCreateInterface( $pEdit_BO_sq, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

   Local $pValue_BO_sq, $oValue_BO_sq
   $oEdit_BO_sq.GetCurrentPattern( $UIA_ValuePatternId, $pValue_BO_sq )
   $oValue_BO_sq = ObjCreateInterface( $pValue_BO_sq, $sIID_IUIAutomationValuePattern, $dtagIUIAutomationValuePattern )
   $oValue_BO_sq.SetValue($BO_TP)

;------------------- automation ------ BO  STOP LOSS --------------

   Local $pBO_SL
   $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "StopLoss:  ", $pBO_SL )

   Local $pText_BOSL, $oText_BOSL
   $oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pBO_SL, $pText_BOSL )
   $oText_BOSL = ObjCreateInterface( $pText_BOSL, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

   Local $pRawView_StopBO, $oRawView_StopBO
   $oUIAutomation.RawViewWalker( $pRawView_StopBO )
   $oRawView_StopBO = ObjCreateInterface( $pRawView_StopBO, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)

   Local $pEdit_stopBO, $oEdit_stopBO
   $oRawView_StopBO.GetNextSiblingElement( $oText_BOSL, $pEdit_stopBO )
   $oEdit_stopBO = ObjCreateInterface( $pEdit_stopBO, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

   Local $pValue_BO_SL, $oValue_BO_SL
   $oEdit_stopBO.GetCurrentPattern( $UIA_ValuePatternId, $pValue_BO_SL )
   $oValue_BO_SL = ObjCreateInterface( $pValue_BO_SL, $sIID_IUIAutomationValuePattern, $dtagIUIAutomationValuePattern )
   $oValue_BO_SL.SetValue($BO_stop)

;------------------- automation ------ BO  Trailing Stop Checkbox--------------

   If $BO_Trail = "1" Then

	  Local $pTrailingCheckbox
	  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "Trailing StopLoss", $pTrailingCheckbox )

	  Local $pCheckBox_Trail, $oCheckBox_Trail
	  $oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pTrailingCheckbox, $pCheckBox_Trail )
	  $oCheckBox_Trail = ObjCreateInterface( $pCheckBox_Trail, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

	  Local $pToggle_trail, $oToggle_trail
	  $oCheckBox_Trail.GetCurrentPattern( $UIA_TogglePatternId, $pToggle_trail )
	  $oToggle_trail = ObjCreateInterface( $pToggle_trail, $sIID_IUIAutomationTogglePattern, $dtagIUIAutomationTogglePattern )
	  $oToggle_trail.Toggle()

;------------------- automation ------ BO  Trailing Stop Value --------------
	  Local $pTrailGap
	  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "Trail Gap  ", $pTrailGap )

	  Local $pText_TrailSL, $oText_TrailSL
	  $oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pTrailGap, $pText_TrailSL )
	  $oText_TrailSL = ObjCreateInterface( $pText_TrailSL, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

	  Local $pRawView_TrailSL, $oRawView_TrailSL
	  $oUIAutomation.RawViewWalker( $pRawView_TrailSL )
	  $oRawView_TrailSL = ObjCreateInterface( $pRawView_TrailSL, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)

	  Local $pEdit_TrailSL, $oEdit_TrailSL
	  $oRawView_TrailSL.GetNextSiblingElement( $oText_TrailSL, $pEdit_TrailSL )
	  $oEdit_TrailSL = ObjCreateInterface( $pEdit_TrailSL, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

	  Local $pValue_TrailSL, $oValue_TrailSL
	  $oEdit_TrailSL.GetCurrentPattern( $UIA_ValuePatternId, $pValue_TrailSL )
	  $oValue_TrailSL = ObjCreateInterface( $pValue_TrailSL, $sIID_IUIAutomationValuePattern, $dtagIUIAutomationValuePattern )
	  $oValue_TrailSL.SetValue($BO_TrailVal)
   EndIf
EndIf
ConsoleWrite( "All Okay" & @CRLF )

;------------------------- SUBMIT Button--------------------------------
Local $pSubmitButton
$oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "btnSubmit", $pSubmitButton )

Local $pButton1, $oButton1
$oUcNewBuySell.FindFirst( $TreeScope_Descendants, $pSubmitButton, $pButton1 )
$oButton1 = ObjCreateInterface( $pButton1, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

Local $pInvokeButton, $oInvokeButton
$oButton1.GetCurrentPattern( $UIA_InvokePatternId, $pInvokeButton )
$oInvokeButton = ObjCreateInterface( $pInvokeButton, $sIID_IUIAutomationInvokePattern, $dtagIUIAutomationInvokePattern )
$oInvokeButton.Invoke()



;------------------------- Confirm SUBMIT Button--------------------------------
if $Order_Confirm = "0" Then

Sleep(500)
Local $pCondiTWS
$oUIAutomation.CreatePropertyCondition( $UIA_ClassNamePropertyId, "#32770", $pCondiTWS )

Local $pWinbox4, $oWinbox4
$oBuyOrderView.FindFirst( $TreeScope_Descendants, $pCondiTWS, $pWinbox4 )
$oWinbox4 = ObjCreateInterface( $pWinbox4, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

Local $pCondiTWS2
$oUIAutomation.CreatePropertyCondition( $UIA_AutomationIdPropertyId, "6", $pCondiTWS2 )

Local $pYesbutton, $oYesbutton
$oWinbox4.FindFirst( $TreeScope_Descendants, $pCondiTWS2, $pYesbutton )
$oYesbutton = ObjCreateInterface( $pYesbutton, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

Local $pInvokePattern1, $oInvokePattern1
$oYesbutton.GetCurrentPattern( $UIA_InvokePatternId, $pInvokePattern1 )
$oInvokePattern1 = ObjCreateInterface( $pInvokePattern1, $sIID_IUIAutomationInvokePattern, $dtagIUIAutomationInvokePattern )
$oInvokePattern1.Invoke()

EndIf

ConsoleWrite( "Order Sent" & @CRLF )