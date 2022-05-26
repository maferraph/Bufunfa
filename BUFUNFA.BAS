Attribute VB_Name = "Geral"
Public Estado As String

' ******************   WORKSPACE ****************
Global WKSP As Workspace 'Abre uma área de trabalho


' ***************   BANCO DE DADOS    ****************
Global BDBUF As Database 'Siscoval


' ******************   TABELAS    ******************
Global BDBUF_TBMOV As Recordset 'Tabela de Movimentos
Global BDBUF_TBCAD As Recordset 'Tabela de Cadastros
Global BDBUF_TBREG As Recordset 'Tabela de Registros


' ******************   CAMPOS    *******************

' Campos da Tabela Movimentos
Global BDBUF_TBMOV_CPNRG As Field
Global BDBUF_TBMOV_CPHIS As Field
Global BDBUF_TBMOV_CPTIP As Field
Global BDBUF_TBMOV_CPIMP As Field
Global BDBUF_TBMOV_CPCON As Field
Global BDBUF_TBMOV_CPDEM As Field
Global BDBUF_TBMOV_CPDPR As Field
Global BDBUF_TBMOV_CPDOC As Field
Global BDBUF_TBMOV_CPVAL As Field
Global BDBUF_TBMOV_CPSAL As Field

' Campos da Tabela Cadastros
Global BDBUF_TBCAD_CPCON As Field
Global BDBUF_TBCAD_CPNMB As Field
Global BDBUF_TBCAD_CPNMA As Field
Global BDBUF_TBCAD_CPNUA As Field
Global BDBUF_TBCAD_CPNUC As Field
Global BDBUF_TBCAD_CPATU As Field

' Campos da Tabela Registros
Global BDBUF_TBREG_CPNRG As Field
Global BDBUF_TBREG_CPCON As Field
Global BDBUF_TBREG_CPDAT As Field
Global BDBUF_TBREG_CPNIN As Field
Global BDBUF_TBREG_CPNFN As Field
