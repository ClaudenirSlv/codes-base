'Variáveis
'---------
Public bd                   As DAO.Database
Public rst                  As DAO.Recordset
Public fd                   As DAO.Field

'Código
'------
Set bd = OpenDatabase("Q:\GROUPS\BR_SC_JGS_WM_ASSISTENCIA_TECNICA\ASSISTENCIA_TECNICA\Pastas particulares\Claudenir\Controle de Equipamentos\DataBaseEQC.0.0.MDB", False, False)

sql = "INSERT INTO tblEquipments (Patrimonio, Num_Metrologia, Marca, Modelo, Descricao, StatusEquipamento, TipoEquipamento, DataCalibracao, ProximaCalibracao, StatusCalibracao, "
sql = sql & "Criado_Por, Data_Adicao) "
sql = sql & "VALUES('" & sPatrimonio & "', '" & sNumMetrologia & "', '" & strMarca & "', '" & strModelo & "','" & strDescricao & "', '" & strStatusEquip & "', '"
sql = sql & strTipoEquip & "', #" & Format(dtDataCalib, "dd/mm/yyyy") & "#, '" & Format(dtProxCalib, "dd/mm/yyyy") & "', '" & strStatusCalib & "', '" & strCriadoPor & "', #" & Format(dtDataAdicao, "dd/mm/yyyy") & "#)"
bd.Execute (sql)
bd.Close