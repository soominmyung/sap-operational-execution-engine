' Soomin PL&STR VBA, SAP to Excel
' A: PL, B: STR, C: by Item, D: By Cont

Function LastRowWithAnyData(targetSheet As Worksheet, targetColumn As String) As Long
    LastRowWithAnyData = targetSheet.Cells(targetSheet.Rows.Count, targetColumn).End(xlUp).Row
End Function

Sub RunProductListQueryWithDynamicRange()
    Dim conn As Object
    Dim rsA As Object
    Dim rsB As Object
    Dim rsC As Object
    Dim rsD As Object
    Dim sqlQueryA As String
    Dim sqlQueryB As String
    Dim sqlQueryC As String
    Dim sqlPartA1 As String
    Dim sqlPartA2 As String
    Dim sqlPartA3 As String
    Dim sqlPartA4 As String
    Dim sqlPartA5 As String
    Dim sqlPartA6 As String
    Dim sqlPartA7 As String
    Dim sqlPartC1 As String
    Dim sqlPartC2 As String
    Dim sqlPartC3 As String
    Dim sqlPartC4 As String
    Dim sqlPartC5 As String
    Dim inactive As String
    Dim targetSheet1 As Worksheet
    Dim targetSheet2 As Worksheet
    Dim targetSheet3 As Worksheet
    Dim minItemCode As String
    Dim maxItemCode As String
    Dim colToDelete As Variant
    Dim currentWorkbook As Workbook
    Dim maxRowA As Long
    Dim maxRowB As Long
    Dim maxRowC As Long
    Dim maxRowD As Long

    Set conn = CreateObject("ADODB.Connection")

    conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=True;" & _
                            "Data Source=<SERVER_NAME_OR_IP>;" & _
                            "Initial Catalog=<DATABASE_NAME>;" & _
                            "User ID=<DB_USER>;" & _
                            "Password=<DB_PASSWORD>;"
    conn.Open

    minItemCodeA = ThisWorkbook.Sheets("PL&STR").Range("AK1").Value
    maxItemCodeA = ThisWorkbook.Sheets("PL&STR").Range("AM1").Value
    inactive = "N"

    If minItemCodeA = "" Then
        minItemCodeA = ""
    End If
    If maxItemCodeA = "" Then
        maxItemCodeA = "999999"
    End If

    sqlPartA1 = "SET DATEFORMAT ymd; " & _
            "declare  @ItemName  as nvarchar(50) " & _
            "declare  @FrgnName  as nvarchar(50) " & _
            "declare  @Supplier  as nvarchar(50) " & _
            "declare  @inactive  as nvarchar(1) " & _
            "DECLARE  @itmsgrpnam as  nvarchar(100) " & _
            "declare  @WareHouse1 as nvarchar(20) " & _
            "declare  @WareHouse2 as nvarchar(20) " & _
            "declare  @WareHouse3 as nvarchar(20) " & _
            "declare  @WareHouse4 as nvarchar(20) " & _
            "declare  @WareHouse5 as nvarchar(20) " & _
            "declare  @WareHouse7 as nvarchar(20) " & _
            "declare  @WareHouse8 as nvarchar(20) " & _
            "declare  @WareHouse9 as nvarchar(20) " & _
            "declare  @WareHouse10 as nvarchar(20) "

    sqlPartA2 = "SET @ItemName  = '' " & _
            "SET @FrgnName  = '' " & _
            "SET @Supplier  = '' " & _
            "SET @inactive  = 'N' " & _
            "SET @itmsgrpnam = '' " & _
            "IF(@WareHouse1 IS NULL or @WareHouse1='') " & _
            "SET  @WareHouse1='WH01' " & _
            "IF(@WareHouse2 IS NULL or @WareHouse2='') " & _
            "SET  @WareHouse2='WH02' " & _
            "SET  @WareHouse3='WH01_RES' " & _
            "SET  @WareHouse4='WH02_RES' " & _
            "SET  @WareHouse5='WH_AC' " & _
            "SET  @WareHouse7='WH03' " & _
            "SET  @WareHouse8='WH04' " & _
            "SET  @WareHouse9='WH05' " & _
            "SET  @WareHouse10='WH06' "

    sqlPartA3 = "SELECT CAST(T0.ItemCode AS INT) AS 'ItemCode', " & _
            "T0.ItemName AS 'ItemName', " & _
            "T0.FrgnName AS 'FrgnName', " & _
            "T0.SWW AS 'Item Size', " & _
            "T0.SalUnitMsr AS 'Sales UoM', " & _
            "T0.NumInSale AS 'Child Qty', " & _
            "T6.FirmName AS 'Maker', " & _
            "T7.CardName AS 'Main Supplier', " & _
            "MAX( CASE WHEN  T4.Whscode='WH01' THEN ((T4.OnHand - T4.IsCommited) / ISNULL(CASE WHEN T0.NumInSale = 0 Then 1 Else T0.NumInSale End ,1))  END) AS [ WH01 AVAILABLE(BOX) ], " & _
            "MAX( CASE WHEN  T4.Whscode='WH01_RES' THEN ((T4.OnHand - T4.IsCommited) / ISNULL(CASE WHEN T0.NumInSale = 0 Then 1 Else T0.NumInSale End ,1))  END) AS [ WH01_RES AVAILABLE(BOX) ], " & _
            "MAX( CASE WHEN  T4.Whscode='WH02' THEN ((T4.OnHand - T4.IsCommited) / ISNULL(CASE WHEN T0.NumInSale = 0 Then 1 Else T0.NumInSale End ,1))  END) AS [ WH02 AVAILABLE(BOX) ], " & _
            "MAX( CASE WHEN  T4.Whscode='WH02_RES' THEN ((T4.OnHand - T4.IsCommited) / ISNULL(CASE WHEN T0.NumInSale = 0 Then 1 Else T0.NumInSale End ,1))  END) AS [ WH02_RES AVAILABLE(BOX) ], " & _
            "MAX( CASE WHEN  T4.Whscode='WH_AC' THEN ((T4.OnHand - T4.IsCommited) / ISNULL(CASE WHEN T0.NumInSale = 0 Then 1 Else T0.NumInSale End ,1))  END) AS [ WH_AC AVAILABLE(BOX) ], " & _
            "MAX( CASE WHEN  T4.Whscode='WH03' THEN ((T4.OnHand - T4.IsCommited) / ISNULL(CASE WHEN T0.NumInSale = 0 Then 1 Else T0.NumInSale End ,1))  END) AS [ WH03 AVAILABLE(BOX) ], " & _
            "MAX( CASE WHEN  T4.Whscode='WH04' THEN ((T4.OnHand - T4.IsCommited) / ISNULL(CASE WHEN T0.NumInSale = 0 Then 1 Else T0.NumInSale End ,1))  END) AS [ WH04 AVAILABLE(BOX) ], " & _
            "MAX( CASE WHEN  T4.Whscode='WH05' THEN ((T4.OnHand - T4.IsCommited) / ISNULL(CASE WHEN T0.NumInSale = 0 Then 1 Else T0.NumInSale End ,1))  END) AS [ WH05 AVAILABLE(BOX) ], "

    sqlPartA4 = "(SELECT TOP 1 BIN.BinCode from OIBQ BQ0 INNER JOIN OBIN BIN ON BIN.AbsEntry=BQ0.BinAbs WHERE ItemCode=T0.ItemCode  and BQ0.OnHandQty>0 ORDER BY BQ0.AbsEntry) as 'First Bin Location'," & _
             "(SELECT TOP 1 bin.bincode from oitw join obin bin on bin.absentry=oitw.dftbinabs and bin.whscode='WH01' where oitw.itemcode =T0.ItemCode ) AS 'Default Bin Location'," & _
             "CASE WHEN (T0.validFor='N' and T0.frozenFor='Y') THEN N'Yes' ELSE '' END  AS 'Inactive'," & _
             "T0.BWEIGHT1                            AS 'Net Weight'," & _
             "T0.SWEIGHT1                            AS 'Gross Weight'," & _
             "CAST(dbo.[FN_FormatNumeric](T0.SVolume, 3) AS FLOAT)                  AS 'Volume'," & _
             "T0.U_Priority                          AS 'Priority'," & _
             "isnull(T0.QryGroup1,'N') as 'Ambient'," & _
             "isnull(T0.QryGroup2,'N') as 'Frozen'," & _
             "isnull(T0.QryGroup3,'N') as 'Chilled'," & _
             "TI.ItmsGrpNam     as 'ItemGroup'," & _
             "T0.DfltWH as 'Default WH'"

    sqlPartA5 = "FROM OITM T0 " & _
             "INNER JOIN OITW T4 ON T0.ItemCode = T4.ItemCode " & _
             "LEFT JOIN RDR1 T9 ON T9.ItemCode = T0.ItemCode " & _
             "LEFT JOIN PKL1 T10 ON T10.OrderLine = T9.LineNum AND T9.DocEntry = T10.OrderEntry " & _
             "LEFT JOIN OPKL T11 ON T10.AbsEntry = T11.AbsEntry " & _
             "LEFT OUTER JOIN ITM9 TC1 ON T0.ItemCode = TC1.ItemCode AND TC1.PriceList = 1 AND TC1.UomEntry = T0.SUoMEntry " & _
             "LEFT OUTER JOIN ITM1 TC ON T0.ItemCode = TC.ItemCode AND TC.PriceList = 1 AND TC.UomEntry = T0.SUoMEntry " & _
             "LEFT OUTER JOIN ITM1 T5 ON T0.ItemCode = T5.ItemCode AND T5.PriceList = 2 " & _
             "LEFT OUTER JOIN OMRC T6 ON T0.FirmCode = T6.FIRMCODE " & _
             "LEFT OUTER JOIN OCRD T7 ON T0.CardCode = T7.CardCode " & _
             "LEFT OUTER JOIN OITB TI ON T0.itmsgrpcod = TI.itmsgrpcod "

    sqlPartA6 = "WHERE ( T0.ItemCode BETWEEN '" & minItemCodeA & "' AND '" & maxItemCodeA & "' ) " & _
             "AND ISNULL(T0.ItemName,'') like CASE WHEN @ItemName = '' THEN ISNULL(T0.ItemName,'') ELSE  @ItemName END " & _
             "AND ISNULL(T0.FrgnName,'') like CASE WHEN @FrgnName = '' THEN ISNULL(T0.FrgnName,'') ELSE  @FrgnName END " & _
             "AND ISNULL(T7.CardName,'') = CASE WHEN @Supplier = '' THEN ISNULL(T7.CardName,'') ELSE @Supplier END " & _
             "AND ISNULL(TI.itmsgrpnam,0) = CASE WHEN @itmsgrpnam = '' THEN ISNULL(TI.itmsgrpnam,'') ELSE @itmsgrpnam END " & _
             "AND ISNULL(T0.FrozenFor,0) = CASE WHEN @inactive = 'Y' THEN 'Y' WHEN @inactive='N' THEN 'N' ELSE ISNULL(T0.FrozenFor,0)  END "

    sqlPartA7 = "GROUP BY T0.ItemCode, T0.ItemName, T0.FrgnName, T0.SWW, T0.SalUnitMsr, T0.NumInSale, T6.FirmName, T7.CardName, " & _
             "T0.validFor, T0.frozenFor, T0.BWeight1, T0.SWeight1, T0.SVolume, T0.U_Priority, " & _
             "T0.QryGroup1, T0.QryGroup2, T0.QryGroup3, TI.ItmsGrpNam, T0.DfltWH " & _
             "ORDER BY T0.ItemCode;"

    sqlQueryA = sqlPartA1 & sqlPartA2 & sqlPartA3 & sqlPartA4 & sqlPartA5 & sqlPartA6 & sqlPartA7

    Set rsA = CreateObject("ADODB.Recordset")
    rsA.Open sqlQueryA, conn, 3, 1

    If rsA.EOF Then
        maxRowA = 0
    Else
        maxRowA = 0
        Do While Not rsA.EOF
            maxRowA = maxRowA + 1
            rsA.MoveNext
        Loop
        rsA.MoveFirst
    End If

    sqlQueryB = "select 'Stock Request' as DocType,T0.docentry, T1.CardCode,T1.CardName, CAST(T0.itemcode AS INT), " & _
             "T0.Quantity,T0.UomCode,T0.OpenCreQty as 'Open Qty', T0.UomCode as 'Open Qty UOM', " & _
             "T0.FromWhsCod, T0.WhsCode as [To Warehouse], T1.[DocDate] as [Posting Date], T1.[DocDueDate] as [Delivery Date], T2.SlpName as 'Account Manager'," & _
             "T3.U_Name as 'Created by' " & _
             "FROM OITM A0 " & _
             "LEFT JOIN WTQ1 T0 on A0.ItemCode=T0.ItemCode " & _
             "LEFT JOIN OWTQ T1 on T0.DocEntry=T1.DocEntry " & _
             "LEFT JOIN OSLP T2 on T1.SlpCode = T2.SlpCode " & _
             "LEFT JOIN OUSR T3 on T3.UserId = T1.UserSign " & _
             "WHERE T0.LineStatus NOT IN ('c','l','r') and ( T0.ItemCode between '" & minItemCodeA & "' AND '" & maxItemCodeA & "' ) AND " & _
             "T0.FromWhsCod= (CASE WHEN ISNULL('','') ='' THEN  T0.FromWhsCod ELSE  '' END)" & _
             "ORDER BY T0.ItemCode;"

    Set rsB = CreateObject("ADODB.Recordset")
    rsB.Open sqlQueryB, conn, 3, 1

    If rsB.EOF Then
        maxRowB = 0
    Else
        maxRowB = 0
        Do While Not rsB.EOF
            maxRowB = maxRowB + 1
            rsB.MoveNext
        Loop
        rsB.MoveFirst
    End If

    Set currentWorkbook = ThisWorkbook
    Set targetSheet1 = currentWorkbook.Sheets("PL&STR")

    lastUsedRowA = LastRowWithAnyData(targetSheet1, "A")
    lastUsedRowB = LastRowWithAnyData(targetSheet1, "AO")

    targetSheet1.Activate
    targetSheet1.Range("F2:AG" & targetSheet1.Rows.Count).ClearContents
    targetSheet1.Range("AS2:BG" & targetSheet1.Rows.Count).ClearContents
    targetSheet1.Range("F2").CopyFromRecordset rsA
    targetSheet1.Range("AS2").CopyFromRecordset rsB

    If lastUsedRowA > maxRowA + 1 Then
        If maxRowA > 0 Then
            targetSheet1.Range("A" & maxRowA + 2 & ":E" & lastUsedRowA).ClearContents
        End If
        If maxRowA = 0 And lastUsedRowA > 2 Then
            targetSheet1.Range("A3" & ":E" & lastUsedRowA).ClearContents
        End If
    End If

    If lastUsedRowB > maxRowB + 1 Then
        If maxRowB > 0 Then
            targetSheet1.Range("AO" & maxRowB + 2 & ":AR" & lastUsedRowB).ClearContents
        End If
        If maxRowB = 0 And lastUsedRowB > 2 Then
            targetSheet1.Range("AO3" & ":AR" & lastUsedRowB).ClearContents
        End If
    End If

    If lastUsedRowA > 1 And lastUsedRowA < maxRowA + 1 Then
        lastFormulaRowA = lastUsedRowA
        targetSheet1.Range("A" & lastFormulaRowA & ":E" & lastFormulaRowA).AutoFill _
            Destination:=targetSheet1.Range("A" & lastFormulaRowA & ":E" & maxRowA + 1)
    End If

    If lastUsedRowB > 1 And lastUsedRowB < maxRowB + 1 Then
        lastFormulaRowB = lastUsedRowB
        targetSheet1.Range("AO" & lastFormulaRowB & ":AR" & lastFormulaRowB).AutoFill _
            Destination:=targetSheet1.Range("AO" & lastFormulaRowB & ":AR" & maxRowB + 1)
    End If

    rsA.Close
    rsB.Close

    sqlPartC1 = "With RankedResults AS (SELECT DENSE_RANK() OVER (ORDER BY a.ResDocNum,a.DocEntry,a.VisOrder,a.U_ATATime DESC) AS Row,* FROM " & _
                "(SELECT DISTINCT OP.DocStatus, CAST(PC.ItemCode AS INT) AS ItemCode, PC.VisOrder,PC.Dscription AS ItemDesc, OITM.SWW AS Size, PC.UomCode AS UoM," & _
                "Ch.DocEntry AS ResDocEntry,Ch.DocNum as ResDocNum,OP.DocEntry,OP.DocNum,CH.CardCode, " & _
                "REPLACE( CH.CardName,'','') CardName,CH.U_ORDNO U_SeqNo ,CH.U_DOCSENT,CH.U_EMAILSENT,CH.U_DOCVIA,CH.U_CNTNO,CH.U_Comments, " & _
                "CH.U_PINum,CH.U_CONTTY,CH.U_CONTSZ, CH.U_ETD  U_ETD, CH.U_ETA " & _
                "U_ETA, CAST(CH.U_ATA as date) U_ATA, CASE WHEN  Len(CH.U_ATATime) =1 then replace(CH.U_ATATime,'0',NULL) " & _
                "ELSE CH.U_ATATime END AS U_ATATime, CH.U_Shipp,CH.U_LSP ,CH.U_POL,CH.U_POD,CH.U_VESSEL,CH.U_SHIPTYPE ,CH.Comments U_Rem, CH.DocTotal as 'AP Res', "

    sqlPartC2 = "OP.DocTotal AS 'PO total',PC.Quantity AS Quantity from OPCH CH INNER JOIN PCH1 PC on PC.DocEntry=CH.DocEntry  AND CH.isIns='Y' AND CH.CANCELED <>'Y' AND CH.CANCELED<>'C' " & _
                "LEFT JOIN OPOR OP ON PC.BaseEntry=OP.DocEntry and PC.BaseType=OP.ObjType and OP.WddStatus NOT IN ('W','C','N') " & _
                "AND OP.CANCELED <>'Y'  LEFT JOIN POR1 ON OP.DocEntry = POR1.DocEntry LEFT JOIN OITM ON OITM.ItemCode=PC.ItemCode " & _
                "WHERE 1= CASE WHEN POR1.TargetType =-1 and OP.DocStatus='C' THEN 2 ELSE 1 END AND CH.U_SHIPTYPE='I' AND PC.Itemcode BETWEEN '" & minItemCodeA & "' AND '" & maxItemCodeA & "' " & _
                "AND ISNULL(PC.TargetType,0) " & _
                "IN(-1,0) and 0=(case when OP.DocType='S' and OP.DocStatus='C' then 1 when OP.DocType<>'S' then 0  else 0 end) and CH.DocEntry " & _
                "NOT IN (SELECT DocEntry FROM PCH1 where DocEntry=CH.DocEntry AND "

    sqlPartC3 = "PCH1.TargetType=20) UNION ALL SELECT DISTINCT  OP.DocStatus ,  POR1.ItemCode,POR1.VisOrder,POR1.Dscription,OITM.SWW, " & _
                "POR1.UomCode,Ch.DocEntry as ResDocEntry,Ch.DocNum as ResDocNum,OP.DocEntry,OP.DocNum,OP.CardCode,OP.CardName,OP.U_ORDNO U_SeqNo, " & _
                "OP.U_DOCSENT,OP.U_EMAILSENT,OP.U_DOCVIA,OP.U_CNTNO,OP.U_Comments,OP.U_PINum," & _
                "OP.U_CONTTY, " & _
                "OP.U_CONTSZ,OP.U_ETD,OP.U_ETA,OP.U_ATA, " & _
                "CAST(OP.U_ATATime as int) as U_ATATime,OP.U_Shipp,OP.U_LSP ,OP.U_POL,OP.U_POD,OP.U_VESSEL,OP.U_SHIPTYPE ,OP.Comments U_Rem, CH.DocTotal as 'AP Res', " & _
                "OP.DocTotal  as 'PO total',POR1.Quantity from OPCH CH INNER JOIN PCH1 PC on PC.DocEntry=CH.DocEntry  AND CH.isIns='Y' AND CH.CANCELED <>'Y' " & _
                "RIGHT JOIN OPOR OP ON PC.BaseEntry=OP.DocEntry and PC.BaseType=OP.ObjType INNER JOIN POR1 ON OP.DocEntry = POR1.DocEntry "

    sqlPartC4 = "INNER JOIN OITM ON OITM.ItemCode=POR1.ItemCode " & _
                "WHERE OP.WddStatus NOT IN ('W','C','N') AND OP.CANCELED <>'Y' AND Ch.DocEntry IS NULL " & _
                "and 1= CASE WHEN POR1.TargetType =-1 and OP.DocStatus='C' then 2 ELSE 1 end AND OP.DocNum > 0  AND " & _
                "OP.U_SHIPTYPE='I' AND POR1.ItemCode BETWEEN '" & minItemCodeA & "' AND '" & maxItemCodeA & "' AND ISNULL(PC.TargetType,0) IN(-1,0) and 0=(CASE WHEN OP.DocType='S' and OP.DocStatus='C' " & _
                "THEN 1 WHEN OP.DocType<>'S' THEN 0 ELSE 0 end) and CH.DocEntry NOT IN (SELECT DocEntry FROM PCH1 WHERE DocEntry=CH.DocEntry AND " & _
                "PCH1.TargetType=20)and OP.DocEntry NOT IN ( select distinct  DocEntry from POR1 where DocEntry = OP.DocEntry and POR1.TargetType = 20 ) ) AS a ) "

    sqlPartC5 = "SELECT ItemCode, ItemDesc, Size, Quantity, UoM, ResDocEntry, ResDocNum, RR.DocEntry, DocNum, U_DOCSENT, U_EMAILSENT, U_CNTNO, " & _
                "RR.CardCode, RR.CardName, U_SeqNo, U_Comments, " & _
                "COALESCE(CON3.Descr, U_CONTTY) AS U_CONTTY_Replaced, " & _
                "U_CONTSZ, U_ETD, U_ETA, U_ATA, U_ATATime, COALESCE(OCRD.CardName, U_LSP) AS U_LSP_Replaced, " & _
                "U_Shipp, U_POL, " & _
                "COALESCE(CON1.Descr, U_POD) AS U_POD_Replaced, " & _
                "U_VESSEL, U_DOCVIA, " & _
                "COALESCE(CON2.Descr, U_SHIPTYPE) AS U_SHIPTYPE_Replaced, " & _
                "U_Rem " & _
                "FROM RankedResults RR " & _
                "LEFT JOIN UFD1 CON1 ON RR.U_POD = CON1.FldValue AND CON1.TableID = '@CONMANCH' " & _
                "LEFT JOIN UFD1 CON2 ON RR.U_SHIPTYPE = CON2.FldValue AND CON2.TableID = '@CONMANCH' " & _
                "LEFT JOIN UFD1 CON3 ON RR.U_CONTTY = CON3.FldValue AND CON3.TableID = '@CONMANCH' " & _
                "LEFT JOIN OCRD ON RR.U_LSP = OCRD.CardCode " & _
                "ORDER BY Row;"

    sqlQueryC = sqlPartC1 & sqlPartC2 & sqlPartC3 & sqlPartC4 & sqlPartC5

    Set rsC = CreateObject("ADODB.Recordset")
    rsC.Open sqlQueryC, conn, 3, 1

    If rsC.EOF Then
        maxRowC = 0
    Else
        maxRowC = 0
        Do While Not rsC.EOF
            maxRowC = maxRowC + 1
            rsC.MoveNext
        Loop
        rsC.MoveFirst
    End If

    Set targetSheet2 = currentWorkbook.Sheets("by Item")
    targetSheet2.Activate

    lastUsedRowC = LastRowWithAnyData(targetSheet2, "G")

    targetSheet2.Range("G2:AK" & targetSheet2.Rows.Count).ClearContents
    targetSheet2.Range("G2").CopyFromRecordset rsC

    If lastUsedRowC > maxRowC + 1 Then
        If maxRowC > 0 Then
            targetSheet2.Range("A" & maxRowC + 2 & ":F" & lastUsedRowC).ClearContents
        End If
        If maxRowC = 0 And lastUsedRowC > 2 Then
            targetSheet2.Range("A3" & ":F" & lastUsedRowC).ClearContents
        End If
    End If

    If lastUsedRowC > 1 And lastUsedRowC < maxRowC + 1 Then
        lastFormulaRowC = lastUsedRowC
        targetSheet2.Range("A" & lastFormulaRowC & ":F" & lastFormulaRowC).AutoFill _
        Destination:=targetSheet2.Range("A" & lastFormulaRowC & ":F" & maxRowC + 1)
    End If

    With targetSheet2
        .Columns(21).NumberFormat = "0"
        .Columns(21).Value = .Columns(21).Value
        .Columns(36).NumberFormat = "0"
        .Columns(36).Value = .Columns(36).Value
    End With

    rsC.Close

    sqlPartD1 = "WITH Results AS (Select ROW_NUMBER( )  OVER (ORDER BY CardName,U_ETA asc) AS Row,* FROM " & _
                "( select Distinct  OP.DocStatus ,  CH.DocEntry as ResDocEntry,'N' " & _
                "as Sel ,CH.DocNum as ResDocNum,OP.DocEntry,OP.DocNum,OP.CardCode, " & _
                "REPLACE( OP.CardName,'''','') CardName,OP.U_ORDNO U_SeqNo ,OP.U_DOCSENT,OP.U_EMAILSENT,OP.U_DOCVIA, " & _
                "OP.U_CNTNO,OP.U_Comments,OP.U_PINum,OP.U_CONTTY,OP.U_CONTSZ,OP.U_ETD U_ETD, " & _
                "OP.U_ETA U_ETA,OP.U_ATA U_ATA, case when  Len(OP.U_ATATime) =1 then " & _
                "REPLACE(OP.U_ATATime,'0',NULL) else OP.U_ATATime end  as U_ATATime, OP.U_Shipp,OP.U_LSP ,OP.U_POL,OP.U_POD,OP.U_VESSEL,OP.U_SHIPTYPE,OP.Comments U_Rem, "

    sqlPartD2 = "CH.DocTotal as 'AP Res', OP.DocTotal as 'PO total' from OPCH CH INNER JOIN PCH1 PC on PC.DocEntry=CH.DocEntry AND CH.isIns='Y' AND CH.CANCELED <>'Y' " & _
                "RIGHT JOIN OPOR OP ON PC.BaseEntry=OP.DocEntry and PC.BaseType=OP.ObjType INNER JOIN POR1 ON OP.DocEntry = POR1.DocEntry  where OP.WddStatus " & _
                "NOT IN ('W','C','N') and OP.CANCELED <>'Y' and 1= case when POR1.TargetType =-1 and OP.DocStatus='C' then 2 else 1 end   AND OP.DocNum > 0 " & _
                "AND OP.U_SHIPTYPE='I'  AND isnull(PC.TargetType,0) IN(-1,0) and 0=(case when OP.DocType='S' and OP.DocStatus='C' " & _
                "THEN 1 when OP.DocType<>'S' then 0  else 0 end) and CH.DocEntry  NOT IN (select DocEntry from PCH1 where DocEntry=CH.DocEntry and  PCH1.TargetType=20) " & _
                "AND OP.DocEntry  NOT IN (select DocEntry from POR1 where DocEntry=OP.DocEntry and  POR1.TargetType=20)) as a) "

    sqlPartD3 = "SELECT Sel, ResDocEntry, ResDocNum, R.DocEntry, DocNum, U_DOCSENT, U_EMAILSENT, U_DOCVIA, U_CNTNO, U_Comments, R.CardCode, R.CardName, U_SeqNo, " & _
                "COALESCE(CONC.Descr, U_CONTTY) AS U_CONTTY_Replaced, " & _
                "U_CONTSZ, U_ETD, U_ETA, U_ATA, U_ATATime, " & _
                "COALESCE(OCRD.CardName, U_LSP) AS U_LSP_Replaced, " & _
                "U_Shipp, U_POL, " & _
                "COALESCE(CONA.Descr, U_POD) AS U_POD_Replaced, " & _
                "U_VESSEL, " & _
                "COALESCE(CONB.Descr, U_SHIPTYPE) AS U_SHIPTYPE_Replaced, " & _
                "U_Rem, [PO total], [AP Res] "

    sqlPartD4 = "FROM Results AS R " & _
                "LEFT JOIN UFD1 CONA ON R.U_POD = CONA.FldValue AND CONA.TableID = '@CONMANCH' " & _
                "LEFT JOIN UFD1 CONB ON R.U_SHIPTYPE = CONB.FldValue AND CONB.TableID = '@CONMANCH' " & _
                "LEFT JOIN UFD1 CONC ON R.U_CONTTY = CONC.FldValue AND CONC.TableID = '@CONMANCH' " & _
                "LEFT JOIN OCRD ON R.U_LSP = OCRD.CardCode " & _
                "ORDER BY DocNum"

    sqlQueryD = sqlPartD1 & sqlPartD2 & sqlPartD3 & sqlPartD4

    Set rsD = CreateObject("ADODB.Recordset")
    rsD.Open sqlQueryD, conn, 3, 1

    Set targetSheet3 = currentWorkbook.Sheets("by cont.")
    targetSheet3.Activate

    targetSheet3.Range("A2:AB" & targetSheet3.Rows.Count).ClearContents
    targetSheet3.Range("A2").CopyFromRecordset rsD

    With targetSheet3
        .Columns(13).NumberFormat = "0"
        .Columns(13).Value = .Columns(13).Value
        .Columns(26).NumberFormat = "0"
        .Columns(26).Value = .Columns(26).Value
    End With

    MsgBox "Data has been updated successfully."
    rsD.Close

    targetSheet1.Activate

    conn.Close

End Sub

