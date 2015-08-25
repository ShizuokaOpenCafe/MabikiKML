'=====================================================================
' KMLファイル間引き処理
'
'　使い方
'　cscript MabikiKML.vbs InputFile OutputFile Threthhold
'　引数
'　　第1引数：入力KMLファイル
'　　第2引数：出力KMLファイル
'　　第3引数：間引き閾値
'　　　　　　10進数の緯度、又は経度いずれかがこの値以下は間引く
'　　　　　　0.00001で約1m 緯度35、経度138付近の場合
'=====================================================================

	'間引き距離閾値
	THRESHHOLD = CDbl(WScript.Arguments(2))

	'----------------------------------------------------------------
	' 入力ファイル
	'----------------------------------------------------------------
	Set oInStream = CreateObject("ADODB.Stream")
	oInStream.type          = 2
	oInStream.charset       = "UTF-8"
	oInStream.LineSeparator = 10
	oInStream.Open
	oInStream.LoadFromFile WScript.Arguments(0)

	'----------------------------------------------------------------
	' 出力ファイル
	'----------------------------------------------------------------
	Set oOutStream = CreateObject("ADODB.Stream")
	oOutStream.type          = 2
	oOutStream.charset       = "UTF-8"
	oOutStream.LineSeparator = 10
	oOutStream.Open

	iTotalCount   = 0
	iPlotCount    = 0
	lastLat       = 0.0
	lastLon       = 0.0
	
	'----------------------------------------------------------------
	'元ファイルオープン
	'----------------------------------------------------------------
	Do While oInStream.EOS = False
	
		'------------------------------------------------------------
		'1行読み込み
		'------------------------------------------------------------
		strLineData = oInStream.ReadText(-2)
		
		'------------------------------------------------------------
		'<coordinates>…</coordinates>の行か判定
		'------------------------------------------------------------
		If Instr( strLineData, "<coordinates>" ) > 0 Then
			isMabikiData = True
		Else
			isMabikiData = False
		End If
		
		'------------------------------------------------------------
		'<coordinates>…</coordinates>の行の場合は間引き処理
		'------------------------------------------------------------
		If isMabikiData = True Then
		
			'--------------------------------------------------------
			'中身取り出し
			'--------------------------------------------------------
			strMotoData = strLineData
			strLineData = Replace( strLineData, vbTab, "" )
			strLineData = Replace( strLineData, "<coordinates>", "" )
			strLineData = Replace( strLineData, "</coordinates>", "" )
			
			'--------------------------------------------------------
			'スペース区切りで配列に格納
			'--------------------------------------------------------
			strMabikiMaeData = strLineData
			aPlotDataArray = Split( strLineData, " " )
			
			'--------------------------------------------------------
			'間引き
			'--------------------------------------------------------
			strMabikiData = ""
			
			IF UBound( aPlotDataArray ) > LBound( aPlotDataArray ) Then
			
				For i = LBound( aPlotDataArray ) To UBound( aPlotDataArray )
				
					iTotalCount = iTotalCount + 1
					
					
					aLatLonHeight = Split(aPlotDataArray(i), ",")
					
					If Abs( CDbl(aLatLonHeight(0)) - CDbl(lastLat) ) > THRESHHOLD Or _
					   Abs( CDbl(aLatLonHeight(1)) - CDbl(lastLon) ) > THRESHHOLD Then
					   
					   
					   iPlotCount = iPlotCount + 1
					   
						If strMabikiData <> "" Then
							strMabikiData = strMabikiData & " " & aPlotDataArray(i)
						Else
							strMabikiData = aPlotDataArray(i)
						End If
						
					   lastLat=CDbl(aLatLonHeight(0))
					   lastLon=CDbl(aLatLonHeight(1))
					   
					   
					End If
				Next
				
			Else
				strMabikiData = aPlotDataArray( UBound( aPlotDataArray ) )
				iTotalCount = iTotalCount + 1
				iPlotCount = iPlotCount + 1
			End If
			
			'--------------------------------------------------------
			'中身差し替え
			'--------------------------------------------------------
			strNewData = Replace( strMotoData, strMabikiMaeData, strMabikiData)

			'Wscript.Echo strNewData
			strWriteLine = strNewData
		Else

			'Wscript.Echo strLineData
			strWriteLine = strLineData

		End If
		
		oOutStream.WriteText strWriteLine, 1
	Loop
	
	oInStream.Close
	
	oOutStream.SaveToFile  WScript.Arguments(1), 2
	oOutStream.Close
