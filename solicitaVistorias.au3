#include <Clipboard.au3>
#include <GUIConstantsEx.au3>
#include <WinAPI.au3>
#include <MsgBoxConstants.au3>
#include <Excel.au3>
#include <WindowsConstants.au3>
#include <StringConstants.au3>
#include <img\ImageSearch.au3>
#include <Sound.au3>
#include <Date.au3>

HotKeySet("{END}", "sair")
Opt("WinTitleMatchMode", 2)

;~ ============================================
;~ =================VARIAVEIS==================
;~ ============================================
Global $X=0
Global $Y=0

Global $i = 2

Global $descricaoVistoria

Global $nomeArquivoEmpresarial
Global $nomeArquivoProposta

Global $sWorkbookEmpresarial
Global $sWorkbookProposta

Global $oExcelEmpresarial
Global $oExcelProposta

Global $oWorkbookEmpresarial
Global $oWorkbookProposta

Global $numLinhas

Global $lmiRouboMaquinas = 0
Global $lmiRouboMercadoriasMateriasPrimas = 0
Global $lmiRouboValoresInterior = 0
Global $lmiRouboValoresMaos = 0
Global $maiorDosRoubos = 0

Global $quantidadeItem, $renovacao, $atividade, $ocupacao, $cnae, $vendaval, $incendio



;~ =============================================
;~ ============== PROCESSAMENTO ================
;~ =============================================

main()

;~ =============================================
;~ =============================================
;~ =============================================



Func main()


Local $sPasswd = InputBox("Verificação de Segurança", "Para continuar digite a senha."&@CRLF&@CRLF&"Se tiver algum problema entre em contato com a equipe da GERID. Ramal 2370.", "", "*")

MsgBox(48,"ATENÇÃO", "Para o Script funcionar corretamente siga os seguintes passos:"&@CRLF&@CRLF&@CRLF&"1º. Coloque em sua Área de Trabalho os arquivos Excel de 'Propostas Pendentes' e de 'Empresarial'."&@CRLF&@CRLF&"2º. O nome dos arquivos deve conter as palavras 'Pendentes de Análise' e 'Empresarial'."&@CRLF&@CRLF&"3º. Os arquivos devem estar abertos.")
Sleep(500)

Local $hI = @HOUR
Local $mI = @MIN


If $sPasswd = "caixa2015" Then

   BlockInput(1)
   If ProcessExists("OUTLOOK.EXE") Then
	  ProcessClose("OUTLOOK.EXE")
   EndIf
   If ProcessExists("LYNC.EXE") Then
	  ProcessClose("LYNC.EXE")
   EndIf
   trataPlanilhas();~executa apenas uma vez
   ConsoleWrite($numLinhas)
   For $j = 1 To $numLinhas Step 1
	  $quantidadeItem = 0
	  $descricaoVistoria = ""
	  analisaPropostas()
	  ConsoleWrite(@CRLF)
   Next
   ajustaColunas($oExcelProposta)
   $oWorkbookProposta.Sheets("Plan2").Select
   $oExcelProposta.Range("F1").Select
   Send("^{SPACE}")
   Sleep(400)
   _Excel_RangeReplace($oWorkbookProposta, Default, Default, "SIENG05S - C/ VISTORIA MENOS 3 ANOS. SERV ", "1")
   _Excel_RangeReplace($oWorkbookProposta, Default, Default, "SIENG05S SOLICITACAO GRAVADA COM SUCESSO NUM     ", "")
   Sleep(500)
   _Excel_RangeReplace($oWorkbookProposta, Default, Default, "SIENG05S - UF+CIDADE NAO ENCONTRADA. VERIFIQUE O NOME CORRETO (1)", "Proposta em Cotação.")

   Sleep(1000)

   Run("notepad.exe")
   Sleep(2000)
   Local $hF = @HOUR
   Local $mF = @MIN
   Send("ARQUIVO DIA "&@MDAY&"/"&@MON&"/"&@YEAR&@CRLF&"Início - "&$hI&":"&$mI&@CR&"Fim - "&$hF&":"&$mF&@CRLF)
   Send($numLinhas&" propostas.")
;~    Send($numLinhas&" Propostas"&@CR&" Solicitadas"&@CR&" <1.000.000 E <100.000"&@CR&" SOBREPOSTA"&@CR&" DECLINAR (SERASA)"&@CR&" UF ERRADA"&@CR&" SOL.MANUAL"&@CR&" Empresa Pública"&@CR&" Restrição Operacional"&@CR&" Aprovação Gerencial")

   Sleep(600)

   Local $aSound = _SoundOpen("WindowsHardwareInsert.wav")
   _SoundPlay($aSound, 1)
   _SoundClose($aSound)
   Sleep(1500)
   $aSound = _SoundOpen("WindowsHardwareInsert.wav")
   _SoundPlay($aSound, 1)
   _SoundClose($aSound)

   BlockInput(0)
EndIf
;~ =============================================
;~ =============================================
;~ =============================================
EndFunc ;~ main



Func analisaPropostas()
   WinActivate("Propost")
   $oWorkbookProposta.Sheets("Plan2").Select
   $oExcelProposta.Range("D"&$i).Select

   Sleep(200)
   Send("^c")
   Sleep(200)

   abreSiesEmissao()

   Local $k = 1
   while not (procuraApenas("img\aguardaColarProposta.bmp"));~~enquanto nao encontrar a imagem...
	  Sleep(100)
	  If $k = 80 Then
		 abreSiesEmissao()
		 $k = 1
	  EndIf
	  $k = $k + 1
   WEnd
   Sleep(1500)

   Send("^v")
   Sleep(600)
   Send("{F10}")
   Sleep(100)

   Local $quebra
   while not (procuraApenas("img\aguardaTelaProposta.bmp")) ;~~ procuraApenas -> RETORNO 0 NAO ENCONTROU O VALOR

	  Sleep(50)
	  $quebra = False

	  If (procuraApenas("img\alertaPropostaNCadastrada.bmp")) Then
		 procuraClica("img\alertaFechaAlerta.bmp",1)
		 Sleep(200)
		 colaNaSituacao(2)		;~~Trata caso de proposta nao cadastrada
		 $quebra = True
	  EndIf

	  If procuraApenas("img\propostaDuplicada.bmp") Then
		 Sleep(200)
		 colaNaSituacao(4)
		 $quebra = True
	  EndIf

	  If ($quebra) Then
		 ExitLoop
	  EndIf

   WEnd
;~    MsgBox(0,"","saiu do loop aguardaTela..")
   Sleep(1000)

   If not $quebra Then ;~~se for false entra
	  procuraClicaX("img\resultadoCrivo.bmp",3,1,100)
	  limpaClipBoard()
	  Sleep(300)
	  Send("^c") ;~~copia o crivo do sies
	  Sleep(300)
	  $resultadoCrivo = ClipGet()
	  Local $aArrayCrivo = FileReadToArray("crivoNaoPodem.txt")
	  Local $analisaAlerta = True

	  If @error Then
		 BlockInput(0)
		 MsgBox($MB_SYSTEMMODAL, "Erro 3", "Erro lendo o arquivo de CRIVO. @error: " & @error) ; An error occurred reading the current script file.
		 Exit
	  Else
		 For $k = 0 To UBound($aArrayCrivo) - 1 ; Loop through the array.
			If($resultadoCrivo = $aArrayCrivo[$k]) Then ;~~verifica os crivos que nao podem passar
			   ClipPut("CRIVO não autorizado")
			   colaNaSituacao(1)
			   $analisaAlerta = False
			   ExitLoop
			EndIf
		 Next
	  EndIf

	  If $analisaAlerta Then ;~~ Analisa Aba Alerta
		 procuraClica("img\abaAlertas.bmp",1)
		 Sleep(1000)

;~ 		 If (procuraApenas("img\restricaoOperacionalRepreLegal.bmp")) Then
;~ 			ClipPut("Restrição Operacional")
;~ 			colaNaSituacao(1)
;~ 		 ElseIf (procuraApenas("img\restricaoOperacionalSegurado.bmp")) Then
;~ 			ClipPut("Restrição Operacional")
;~ 			colaNaSituacao(1)
		 If (procuraApenas("img\aprovacaoGerencial.bmp")) Then
			ClipPut("Aprovação Gerencial")
			colaNaSituacao(1)
		 ElseIf (procuraApenas("img\empresaPublica.bmp")) Then
			ClipPut("Empresa Pública")
			colaNaSituacao(1)
;~ 		 ElseIf (procuraApenas("img\propostaSobreposta.bmp")) Then
;~ 			ClipPut("SOBREPOSTA")
;~ 			colaNaSituacao(1)
		 Else
;~ 			ToolTip("Não é caso de 'Proposta Sobreposta'", 30, 30)
;~ 			Sleep(800) ; Sleep to give tooltip time to display
			procuraClica("img\abaLocalRisco.bmp",1)
			Sleep(300)


			limpaClipBoard()
			Sleep(500)
			procuraClicaX("img\quantidade.bmp",3,1,70)
			Sleep(500)
			Send("^c")
			Sleep(400)
			$quantidadeItem = ClipGet()

			limpaClipBoard()
			Sleep(500)
			procuraClicaX("img\renovacao.bmp",1,1,100)
			Sleep(500)
			Send("^c")
			Sleep(400)
			Send("{ESC}")
			Sleep(600)
			$renovacao = ClipGet()


			procuraClica("img\btnCaracteristicaRisco.bmp",1)
			Sleep(300)

			Local $analisaCNAE = True
			Local $analisaCobertura = False

			limpaClipBoard()
			Sleep(500)
			procuraClicaX("img\atividade.bmp",1,1,100)
			Sleep(500)
			Send("^c")
			Sleep(400)
			Send("{ESC}")
			Sleep(600)
			$atividade = ClipGet()

			limpaClipBoard()
			Sleep(500)
			procuraClicaX("img\ocupacao.bmp",1,1,100)
			Sleep(500)
			Send("^c")
			Sleep(400)
			Send("{ESC}")
			Sleep(600)
			$ocupacao = ClipGet()

			limpaClipBoard()
			Sleep(500)
			procuraClicaX("img\cnae.bmp",3,1,100)
			Sleep(900)
			Send("^c")
			Sleep(800)
			$cnae = ClipGet()
			Sleep(300)
			Send("{ESC}")
			Sleep(600)


			If(procuraApenas("img\desocupadosNaoAceitar.bmp")) Then
			   ConsoleWrite("Desocupados, com a clausula particular 104. RISCO SEM ACEITACAO.")
			   ClipPut("Desocupados, com a clausula particular 104. RISCO SEM ACEITACAO.")
			   Sleep(600)
			   colaNaSituacao(1)
			   $analisaCNAE = False
			EndIf

			If $analisaCNAE Then

			   procuraClicaX("img\cnae.bmp",3,1,100)
			   limpaClipBoard()
			   Sleep(500)
			   Send("^c") ;~~copia o CNAE do sies
			   Sleep(200)
			   Local $cnaeSies = ClipGet()
			   Local $aArrayCnae = FileReadToArray("cnaeNaoPodem.txt")
			   $analisaCobertura = True

			   If @error Then
				  BlockInput(0)
				  MsgBox($MB_SYSTEMMODAL, "Erro 3", "Erro lendo o arquivo de CNAE. @error: " & @error) ; An error occurred reading the current script file.
				  Exit
			   Else
				  For $k = 0 To UBound($aArrayCnae) - 1 ; Loop through the array.
					 If($cnaeSies = $aArrayCnae[$k]) Then
   ;~ 					 MsgBox($MB_SYSTEMMODAL, "", $aArrayCnae[$k]) ; Display the contents of the array.
						colaNaSituacao(1)
						$analisaCobertura = False
						ExitLoop
					 EndIf
				  Next
			   EndIf

			EndIf

			If $analisaCobertura Then
			   procuraClica("img\btnCobertura.bmp",1)
			   Sleep(200)
			   Send("{down 5}")
			   Sleep(900)

			   limpaClipBoard()
			   Sleep(500)
			   If (procuraApenas("img\vendavalFumacaQueda.bmp")) Then
				  procuraClicaX("img\vendavalFumacaQueda.bmp",3,1,580)
				  Sleep(500)
				  Send("^c")
				  Sleep(400)
				  Send("{ESC}")
				  Sleep(600)
				  $vendaval = ClipGet()
			   Else
				  $vendaval = 0
			   EndIf


			   If (procuraApenas("img\incendioRaioExplosao.bmp")) Then ;~~ procuraApenas -> RETORNO 0 NAO ENCONTROU O VALOR
				  procuraClicaX("img\incendioRaioExplosao.bmp",3,1,580)
				  Sleep(200)
				  Send("^c")
				  Sleep(200)
				  $incendio = ClipGet()
				  Local $lmiIncendio = Number(StringReplace(ClipGet(),".",""),3)

				  If $lmiIncendio > 1000000 Then
					 If (procuraApenas("img\lixeiraBarraRolagem.bmp")) Then ;~~ procuraApenas -> RETORNO 0 NAO ENCONTROU O VALOR
						procuraRoubos()
						Sleep(200)
						procuraClicaY("img\lixeiraBarraRolagem.bmp",1,1,20)
						Sleep(300)
						Send("{DOWN 5}")
						Sleep(800)

						procuraRoubos()
						Sleep(200)
						procuraClicaX("img\lixeiraBarraRolagem2.bmp",1,1,17)
						Sleep(300)
						Send("{DOWN 5}")
						Sleep(800)

						procuraRoubos()
						$maiorDosRoubos = 0
						$maiorDosRoubos = calculaMaiorRoubo() ;~~esse roubo serve apenas para colocar na planilha
					 Else
						procuraRoubos()
						$maiorDosRoubos = 0
						$maiorDosRoubos = calculaMaiorRoubo() ;~~esse roubo serve apenas para colocar na planilha
					 EndIf

					 limpaClipBoard()

					 solicitaVistoria()
					 colaNaSituacao(1)

				  Else;~ <= 1000000
					 limpaClipBoard()
					 Sleep(100)
					 ClipPut('<1.000.000')
					 Sleep(100)
					 colaNaSituacao(1)
;~ 					 If (procuraApenas("img\lixeiraBarraRolagem.bmp")) Then ;~~ procuraApenas -> RETORNO 0 NAO ENCONTROU O VALOR

;~ 						procuraRoubos()
;~ 						Sleep(200)
;~ 						procuraClicaY("img\lixeiraBarraRolagem.bmp",1,1,20)
;~ 						Sleep(300)
;~ 						Send("{DOWN 5}")
;~ 						Sleep(800)

;~ 						procuraRoubos()
;~ 						Sleep(200)
;~ 						procuraClicaX("img\lixeiraBarraRolagem2.bmp",1,1,17)
;~ 						Sleep(300)
;~ 						Send("{DOWN 5}")
;~ 						Sleep(800)

;~ 						procuraRoubos()
;~ 						$maiorDosRoubos = 0
;~ 						$maiorDosRoubos = calculaMaiorRoubo()

;~ 						If $maiorDosRoubos > 100000 Then
;~ 						   limpaClipBoard()
;~ 						   Sleep(100)
;~ 						   solicitaVistoria()
;~ 						Else
;~ 						   limpaClipBoard()
;~ 						   Sleep(100)
;~ 						   ClipPut('<1.000.000 E <100.000')
;~ 						EndIf

;~ 						Sleep(100)
;~ 						colaNaSituacao(1)

;~ 					 Else

;~ 						procuraRoubos()
;~ 						$maiorDosRoubos = 0
;~ 						$maiorDosRoubos = calculaMaiorRoubo()

;~ 						ConsoleWrite("maiorDosRoubos"&$maiorDosRoubos&@CRLF)

;~ 						If $maiorDosRoubos > 100000 Then
;~ 						   limpaClipBoard()
;~ 						   Sleep(100)
;~ 						   solicitaVistoria()
;~ 						Else
;~ 						   limpaClipBoard()
;~ 						   Sleep(100)
;~ 						   ClipPut('<1.000.000 E <100.000')
;~ 						EndIf
;~ 						sleep(100)
;~ 						colaNaSituacao(1)
;~ 					 EndIf
				  EndIf
			   EndIf
			Else
;~ 			   ToolTip("Avaliar cobertura...", 30, 30)
;~ 			   Sleep(800) ; Sleep to give tooltip time to display
;~ 			   colaNaSituacao(3)
			EndIf
		 EndIf
	  EndIf
   EndIf
EndFunc

Func calculaMaiorRoubo()

   Local $A = ($lmiRouboMaquinas + $lmiRouboMercadoriasMateriasPrimas + Abs($lmiRouboMaquinas - $lmiRouboMercadoriasMateriasPrimas)) / 2

   Local $B = ($A + $lmiRouboValoresInterior + Abs($A - $lmiRouboValoresInterior)) / 2

   Local $maior = ($B + $lmiRouboValoresMaos + Abs($B - $lmiRouboValoresMaos)) / 2
   ConsoleWrite("| maior é= "&$maior)

   Return $maior
EndFunc

Func solicitaVistoria()

   procuraClicaX("img\lixeira.bmp",1,1,30)
   Send("{UP 5}")
   Sleep(1000)
   procuraClica("img\btnVistoria.bmp",1)
   Sleep(300)

   procuraClica("img\solicitarVistoria.bmp",2)

   Local $saida = 1
   While not procuraApenas("img\alertaFechar.bmp")     ;~~RETORNO 0 NAO ENCONTROU O VALOR
	  Sleep(5)

	  if $saida > 200 Then
		 BlockInput(0)

		 Sleep(300)

		 Local $aSound = _SoundOpen("WindowsHardwareInsert.wav")
		 _SoundPlay($aSound, 1)
		 _SoundClose($aSound)
		 _SoundPlay($aSound, 1)
		 _SoundClose($aSound)
		 MsgBox(0,"Script Pausado 1","O SIES demorou para retornar a solicitação, pressione OK para continuar o Script.",5)

		 Sleep(500)
		 BlockInput(1)
		 Sleep(500)
		 ExitLoop
	  EndIf

	  $saida = $saida + 1

   WEnd

   Sleep(500)
   limpaClipBoard()
   procuraClicaX("img\descricaoVistoria.bmp",3,1,134)
   Sleep(300)
   Send("^c")
   Sleep(500)

   $descricaoVistoria = ClipGet()
   Sleep(500)
   If $descricaoVistoria = "SIENG05S - UF+CIDADE NAO ENCONTRADA. VERIFIQUE O NOME CORRETO (1)" Then

	  Global $continua = True
	  editaEndereco()
	  Sleep(200)
	  If $continua Then
		 Sleep(1000)
		 Local $cont = 0

		 While not procuraApenas("img\operacaoEfetuada.bmp")
			Sleep(10)
			If $cont > 80 Then
			   ExitLoop
			EndIf
			$cont = $cont + 1
		 WEnd
		 Sleep(300)

		 procuraClica("img\abaProposta.bmp",1)
		 Sleep(900)
		 procuraClica("img\btnSegurado.bmp",1)

		 Sleep(1000)


		 editaEndereco()

		 Sleep(1000)

		 $cont = 0

		 While not procuraApenas("img\operacaoEfetuada.bmp")
			Sleep(10)
			If $cont > 50 Then
			   ExitLoop
			EndIf
			$cont = $cont + 1
		 WEnd

		 procuraClica("img\abaLocalRisco.bmp",1)
		 Sleep(1000)
		 procuraClica("img\solicitarVistoria.bmp",2)
		 Sleep(1000)
		 Local $saida = 1
		 While not procuraApenas("img\alertaFechar.bmp")     ;~~RETORNO 0 NAO ENCONTROU O VALOR
			Sleep(10)

			if $saida > 200 Then
			   BlockInput(0)

			   Sleep(300)

			   Local $aSound = _SoundOpen("WindowsHardwareInsert.wav")
			   _SoundPlay($aSound, 1)
			   _SoundClose($aSound)
			   _SoundPlay($aSound, 1)
			   _SoundClose($aSound)
			   MsgBox(0,"Script Pausado 2","O SIES demorou para retornar a solicitação, pressione OK para continuar o Script.",5)

			   Sleep(500)
			   BlockInput(1)
			   Sleep(500)
			   ExitLoop
			EndIf
			$saida = $saida + 1
		 WEnd

		 Sleep(500)
		 limpaClipBoard()
		 procuraClicaX("img\descricaoVistoria.bmp",3,1,134)
		 Sleep(300)
		 Send("^c")

		 Sleep(800)
		 $descricaoVistoria = ClipGet()
		 Sleep(500)
		 Sleep(1000)

	  EndIf
   EndIf
EndFunc

Func editaEndereco()
	  limpaClipBoard()
	  Sleep(600)
	  procuraClicaX("img\cep.bmp",3,1,80)
	  Sleep(600)
	  Send("^c")
	  Sleep(500)
	  Local $CEP = ClipGet()

	  $continua = True

	  If ($CEP = "") Then
		 MsgBox(0,"Erro 7","O Cep não foi encontrado")
		 $continua = False
	  EndIf

	  ConsoleWrite($CEP&@CRLF)

	  If $continua = True Then
		 limpaClipBoard()
		 procuraClicaX("img\endereco.bmp",3,1,80)
		 Sleep(500)
		 Send("^c")
		 Sleep(500)
		 Local $endereco = ClipGet()

		 limpaClipBoard()
		 procuraClicaX("img\complemento.bmp",3,1,80)
		 Sleep(500)
		 Send("^c")
		 Sleep(500)
		 Local $complemento = ClipGet()

		 limpaClipBoard()
		 procuraClicaX("img\bairro.bmp",3,1,80)
		 Sleep(500)
		 Send("^c")
		 Sleep(500)
		 Local $bairro = ClipGet()

		 Sleep(500)
		 procuraClica("img\buscaEndereco.bmp",1)
		 Sleep(500)

		 While Not procuraApenas("img\pesquisaEndereco.bmp")
			Sleep(10)
		 WEnd

		 limpaClipBoard()
		 Sleep(100)
		 ClipPut($CEP)
		 Sleep(100)
		 procuraClicaX("img\cepPesquisa.bmp",1,1,76)
		 Sleep(500)
		 Send("^v")
		 Sleep(1000)
		 procuraClica("img\pesquisaEndereco.bmp",1)
		 MouseMove($X-50,$Y)

		 Local $j = 1

		 While Not procuraApenas("img\tickEndereco.bmp")
			Sleep(30)

			If $j > 50  Then
			   $continua = False
			   ExitLoop
			EndIf
			$j = $j + 1
		 WEnd

		 If $continua Then
			procuraClica("img\tickEndereco.bmp",1)

			While procuraApenas("img\pesquisaEndereco.bmp")
			   Sleep(20)
			WEnd
			Sleep(3000)

			limpaClipBoard()
			Sleep(500)
			ClipPut($endereco)
			Sleep(300)
			procuraClicaX("img\endereco.bmp",3,1,80)
			Sleep(400)
   ;~ 		 Send("{BACKSPACE}")
			Send("^v")
			Sleep(1000)


			limpaClipBoard()
			ClipPut($complemento)
			Sleep(500)
			procuraClicaX("img\complemento.bmp",3,1,80)
			Sleep(100)
   ;~ 		 Send("{BACKSPACE}")
			Send("^v")
			Sleep(1000)

			limpaClipBoard()
			ClipPut($bairro)
			Sleep(500)
			procuraClicaX("img\bairro.bmp",3,1,80)
			Sleep(100)
   ;~ 		 Send("{BACKSPACE}")
			Send("^v")
			Sleep(1000)

			Sleep(300)
			Send("{F11}")
		 Else
			ConsoleWrite("Tick nao encontrado - caiu no else "&@CRLF)
			procuraClicaX("img\pesquisaEndereco.bmp",1,2,100)
			Sleep(200)
			Send("{UP}")
			Sleep(600)
			procuraClica("img\alertaFechar.bmp",1)

			While procuraApenas("img\alertaFechar.bmp")
			   Sleep(20)
			WEnd
			Sleep(2000)

			limpaClipBoard()
			procuraClicaX("img\descricaoVistoria.bmp",3,1,134)
			Sleep(600)
			Send("^c")
			Sleep(1000)
			$descricaoVistoria = ClipGet()
			ConsoleWrite($descricaoVistoria)
			Sleep(300)
		 EndIf

		 ConsoleWrite("CEP "&$CEP&@CRLF)
		 ConsoleWrite("endereco "&$endereco&@CRLF)
		 ConsoleWrite("complemento "&$complemento&@CRLF)
		 ConsoleWrite("bairro "&$bairro&@CRLF)
	  EndIf

EndFunc

Func procuraRoubos()

   If (procuraApenas("img\rouboMaquinasMoveisUten.bmp")) Then ;~~ procuraApenas -> RETORNO 0 NAO ENCONTROU O VALOR
	  limpaClipBoard()
	  procuraClicaX("img\rouboMaquinasMoveisUten.bmp",3,1,580)
	  Sleep(200)
	  Send("^c")
	  Sleep(200)
	  $lmiRouboMaquinas = Number(StringReplace(ClipGet(),".",""),3)
	  ConsoleWrite($lmiRouboMaquinas)
   Else
	  $lmiRouboMaquinas = 0
   EndIf

   If (procuraApenas("img\rouboMercadoriasMateriasPrimas.bmp")) Then ;~~ procuraApenas -> RETORNO 0 NAO ENCONTROU O VALOR
	  limpaClipBoard()
	  procuraClicaX("img\rouboMercadoriasMateriasPrimas.bmp",3,1,580)
	  Sleep(200)
	  Send("^c")
	  Sleep(200)
	  $lmiRouboMercadoriasMateriasPrimas = Number(StringReplace(ClipGet(),".",""),3)
	  ConsoleWrite($lmiRouboMercadoriasMateriasPrimas)
   Else
	  $lmiRouboMercadoriasMateriasPrimas = 0
   EndIf

   If (procuraApenas("img\rouboValoresInteriorEstabelecimento.bmp")) Then ;~~ procuraApenas -> RETORNO 0 NAO ENCONTROU O VALOR
	  limpaClipBoard()
	  procuraClicaX("img\rouboValoresInteriorEstabelecimento.bmp",3,1,580)
	  Sleep(200)
	  Send("^c")
	  Sleep(200)
	  $lmiRouboValoresInterior = Number(StringReplace(ClipGet(),".",""),3)
	  ConsoleWrite($lmiRouboValoresInterior)
   Else
	  $lmiRouboValoresInterior = 0
   EndIf

   If (procuraApenas("img\rouboValoresMaosPortadores.bmp")) Then ;~~ procuraApenas -> RETORNO 0 NAO ENCONTROU O VALOR
	  limpaClipBoard()
	  procuraClicaX("img\rouboValoresMaosPortadores.bmp",3,1,580)
	  Sleep(200)
	  Send("^c")
	  Sleep(200)
	  $lmiRouboValoresMaos = Number(StringReplace(ClipGet(),".",""),3)
	  ConsoleWrite($lmiRouboValoresMaos)
   Else
	  $lmiRouboValoresMaos = 0
   EndIf
endfunc

Func trataPlanilhas()
   verificaArquivos()
   editaArquivoPropostasPendentes()
   editaArquivoEmpresarial()
   Send("{ALTDOWN}{TAB}{ALTUP}")
   Sleep(500)
   Send("!{F4}")
   Sleep(1000)
   WinActivate($nomeArquivoProposta)
EndFunc

Func editaArquivoPropostasPendentes()
   limpaClipBoard()


   $arqu = WinActivate("Pendentes de An")
   Sleep(900)
   WinMove($arqu,"",0,0,@DesktopWidth,@DesktopHeight-40)
   Sleep(700)
   Send("{ALTDOWN}{SPACE}{ALTUP}")
   Sleep(300)
   Send("x")


   Sleep(1500)
   Send("{ESC}")
   Sleep(100)
   Send("{LALT}")
   Sleep(500)
   Send("a")
   Sleep(400)
   Send("j")
   Sleep(400)
   Send("i")
   Sleep(400)
   Send("c")
   Sleep(400)
   Local $caminhoArquivo = ClipGet()
   Local $cutted = StringSplit($caminhoArquivo,'///',$STR_ENTIRESPLIT)

;~ 	MsgBox(0,"caminho arquivo proposta",$caminhoArquivo)


   $oExcelProposta = _Excel_Open()

;~    MsgBox(0,"instancia excel Propost",$oExcelProposta)

   $sWorkbookProposta = StringReplace($cutted[2],"%20"," ")
   ConsoleWrite($sWorkbookProposta&@CRLF)
   $oWorkbookProposta = _Excel_BookOpen($oExcelProposta, $sWorkbookProposta, Default, Default, True)

   Local $arquivo = StringSplit($sWorkbookProposta,'\',$STR_ENTIRESPLIT)
   ConsoleWrite($arquivo[$arquivo[0]])
   $nomeArquivoProposta = $arquivo[$arquivo[0]]



   $arqu = WinActivate("Pendentes de An")
   WinMove($arqu,"",0,0,@DesktopWidth,@DesktopHeight-40)
   Sleep(200)
   Send("{ALTDOWN}{SPACE}{ALTUP}")
   Sleep(200)
   Send("x")
   Sleep(600)

   $oWorkbookProposta.Sheets("Dinâmica").Select


   $oExcelProposta.Range("n14").Select
   Sleep(500)
   procuraClica("img\excelSelecionado.bmp",2)
   $oWorkbookProposta.Sheets("Plan1").Select
   $oExcelProposta.Range("b4").Select
   $oWorkbookProposta.Activesheet.Columns("B:B").NumberFormat = "###"

   For $i = 1 To 6 Step 1
	  $oExcelProposta.Range("a4").Select

	  Send("{SHIFTDOWN}{F10}{SHIFTUP}")
	  Send("i")
	  Sleep(500)
	  Send("t")
   Next

   $oExcelProposta.Range("I4").Select;~ contrato
   Send("{CTRLDOWN}{SPACE}{SPACE}{CTRLUP}")
   Sleep(500)
   Send("^x")
   Sleep(500)
   $oExcelProposta.Range("A1").Select
   Sleep(100)
   Send("^v")
   Sleep(300)

   $oExcelProposta.Range("L4").Select;~ ini vigencia
   Send("{CTRLDOWN}{SPACE}{SPACE}{CTRLUP}")
   Sleep(500)
   Send("^x")
   Sleep(500)
   $oExcelProposta.Range("B1").Select
   Sleep(100)
   Send("^v")
   Sleep(300)

   $oExcelProposta.Range("J4").Select;~ produto
   Send("{CTRLDOWN}{SPACE}{SPACE}{CTRLUP}")
   Sleep(500)
   Send("^x")
   Sleep(500)
   $oExcelProposta.Range("C1").Select
   Sleep(100)
   Send("^v")
   Sleep(300)

   $oExcelProposta.Range("H4").Select;~ proposta
   Send("{CTRLDOWN}{SPACE}{SPACE}{CTRLUP}")
   Sleep(500)
   Send("^x")
   Sleep(500)
   $oExcelProposta.Range("D1").Select
   Sleep(100)
   Send("^v")
   Sleep(300)

   $oExcelProposta.Range("K4").Select;~ CPF
   Send("{CTRLDOWN}{SPACE}{SPACE}{CTRLUP}")
   Sleep(500)
   Send("^x")
   Sleep(500)
   $oExcelProposta.Range("E1").Select
   Sleep(100)
   Send("^v")
   Sleep(300)

   For $i = 1 To 27 Step 1
	  $oExcelProposta.Columns("G:G").EntireColumn.DELETE
   Next

   ajustaColunas($oExcelProposta)

EndFunc

Func editaArquivoEmpresarial()
   limpaClipBoard()
   $arqu = WinActivate("Empresari")
   Sleep(900)
   WinMove($arqu,"",0,0,@DesktopWidth,@DesktopHeight-40)
   Sleep(700)
   Send("{ALTDOWN}{SPACE}{ALTUP}")
   Sleep(300)
   Send("x")

   Sleep(800)
   Send("{ESC}")
   Send("{LALT}")
   Sleep(500)
   Send("a")
   Sleep(400)
   Send("j")
   Sleep(400)
   Send("i")
   Sleep(400)
   Send("c")
   Sleep(400)
   Local $caminhoArquivo = ClipGet()
   Local $cutted = StringSplit($caminhoArquivo,'///',$STR_ENTIRESPLIT)

   $oExcelEmpresarial = _Excel_Open()

   $sWorkbookEmpresarial = StringReplace($cutted[2],"%20"," ")
;~    ConsoleWrite($sWorkbookEmpresarial&@CRLF)
   $oWorkbookEmpresarial = _Excel_BookOpen($oExcelEmpresarial, $sWorkbookEmpresarial, Default, Default, True)

   $arqu = WinActivate("Empresari")
   Sleep(200)
   WinMove($arqu,"",0,0,@DesktopWidth,@DesktopHeight-40)
   Sleep(200)
   Send("{ALTDOWN}{SPACE}{ALTUP}")
   Sleep(200)
   Send("x")

   $oWorkbookEmpresarial.Sheets("Plan1").Select
   $oExcelEmpresarial.Range("A1").Select
   Send("^+"&"{right}")
   Sleep(600)
   Send("!s")
   Sleep(600)
   If _ImageSearch("img\filtroExcel.bmp",0,$X,$Y,0) = 1 Then
	  Send("!f")
   EndIf

   Local $arquivo = StringSplit($sWorkbookEmpresarial,'\',$STR_ENTIRESPLIT)
;~    ConsoleWrite($arquivo[$arquivo[0]])
   $nomeArquivoEmpresarial = $arquivo[$arquivo[0]]

;~    Local $oWorkbookProposta = _Excel_BookOpen($oExcelProposta, $sWorkbookProposta, Default, Default, True)

   Sleep(700)
   Send("{ALTDOWN}{TAB}{ALTUP}")
   Sleep(300)
   Send("{RIGHT}{DOWN}")
   Send("{ESC}")
   Sleep(200)


ConsoleWrite("= CONT.SE('["&$nomeArquivoEmpresarial&"]Plan1'!$F:$F;[@PropostaRen])")
   Send("= CONT.SE('["&$nomeArquivoEmpresarial&"]")
   Send("Plan1'!$F:$F;[@PropostaRen])",1)
   Send("{ENTER}")
   Send("{ESC}")

   $oExcelProposta.Range("f1").Select
   Sleep(500)
   If procuraClica("img\excelSelecionado.bmp",1) = 1 Then
	  If procuraClica("img\excelSelecionado3.bmp",1) = 1 Then
		 BlockInput(0)
		 MsgBox(0,"Erro P5","Função procuraClica não retornou valor.")
		 Exit
	  EndIf
   EndIf
   Sleep(800)
   Send("{TAB 7}")
   Sleep(800)
   Send("0")
   Sleep(2500)

   If _ImageSearch("img\nenhumaCorrespondencia.bmp",0,$X,$Y,0) = 1 Then
	  ;~~nenhuma correspondencia
	  Send("{ESC 2}")
	  Sleep(200)
	  Send("{ALTDOWN}{TAB}{ALTUP}")

	  MsgBox(0,"Erro","Nenhuma proposta encontrada.")
	  Exit
;~ 	  filtroDeCores()
   Else
	  limpaClipBoard()
	  Sleep(200)
	  Send("{ENTER}")
	  Sleep(100)
	  Send("{DOWN}")
	  Send("{LEFT}")
	  Send("{CTRLDOWN}{SHIFTDOWN}{DOWN}{LEFT}{SHIFTUP}{CTRLUP}")
	  Sleep(150)
	  Send("^c")
	  Sleep(300)

	  _Excel_SheetAdd($oWorkbookProposta, -1, False, 1, "Plan2")
	  Sleep(1000)
	  $oWorkbookEmpresarial.Sheets("Plan2").Select

	  $oExcelProposta.range("A1").FormulaR1C1 = "Contrato"
	  $oExcelProposta.range("B1").FormulaR1C1 = "InicioVigencia"
	  $oExcelProposta.range("C1").FormulaR1C1 = "Produtos"
	  $oExcelProposta.range("D1").FormulaR1C1 = "PropostaRen"
	  $oExcelProposta.range("E1").FormulaR1C1 = "CPF/CNPJ"
	  $oExcelProposta.range("F1").FormulaR1C1 = "Situação"
	  $oExcelProposta.range("G1").FormulaR1C1 = "QT_ItemLocalRisco"
	  $oExcelProposta.range("H1").FormulaR1C1 = "Seguro"
	  $oExcelProposta.range("I1").FormulaR1C1 = "Atividade"
	  $oExcelProposta.range("J1").FormulaR1C1 = "Ocupacao"
	  $oExcelProposta.range("K1").FormulaR1C1 = "CNAE"
	  $oExcelProposta.range("L1").FormulaR1C1 = "IS"
	  $oExcelProposta.range("M1").FormulaR1C1 = "Vendaval"
	  $oExcelProposta.range("N1").FormulaR1C1 = "Roubo"
	  $oExcelProposta.range("A2").select
	  Send("^v")
	  Sleep(600)
	  ajustaColunas($oExcelProposta)
	  Sleep(600)

	  $oExcelProposta.range("F2").select
	  Sleep(200)
	  Send("= SUBTOTAL(3;A2:A500){ENTER}")
	  Sleep(400)
	  $oExcelProposta.range("F2").Select
	  Send("^c")
	  Sleep(400)
	  $numLinhas = ClipGet()
	  Sleep(300)
	  $oExcelProposta.range("F2").Select
	  Send("{DEL}")
	  Sleep(300)

;~ =----------------------------------
;~ 	  trata os casos que ja estao ok
;~ 	  $oExcelProposta.Sheets("Plan1").Select
;~ 	  Sleep(700)
;~ 	  $oExcelProposta.Range("f1").Select
;~ 	  Sleep(800)
;~ 	  procuraClica("img\excelSelecionado.bmp",1)
;~ 	  Send("{TAB 7}")
;~ 	  Sleep(200)
;~ 	  Send("1")
;~ 	  Sleep(2500)

;~ 	  limpaClipBoard()
;~ 	  Sleep(200)
;~ 	  Send("{ENTER}")
;~ 	  Sleep(100)
;~ 	  Send("{DOWN}")
;~ 	  Send("{LEFT}")
;~ 	  Send("{CTRLDOWN}{SHIFTDOWN}{DOWN}{LEFT}{SHIFTUP}{CTRLUP}")
;~ 	  Sleep(150)
;~ 	  Send("^c")
;~ 	  Sleep(1000)

;~ 	  $oExcelProposta.Sheets("Plan2").Select
;~ 	  Sleep(600)

;~ 	  $oExcelProposta.range("A2").select
;~ 	  Sleep(300)
;~ 	  Send("^{DOWN}")
;~ 	  Sleep(200)
;~ 	  Send("{DOWN 2}")
;~ 	  Sleep(100)
;~ 	  Send("^v")
;~ 	  Sleep(600)
;~ 	  Send("^{RIGHT}")
;~ 	  Sleep(300)
;~ 	  Send("{RIGHT}")
;~ 	  Sleep(600)


;~ 	  Send("= PROCV($D"&String(Int($numLinhas) + 3)&";'"&$nomeArquivoEmpresarial&"'!Tabela2[[Proposta]:[Nº Solic Serv]];3)",1)
;~ 	  Send("{ENTER}")
;~ 	  Send("{ESC}")

;~ 	  $oExcelProposta.range("F"&String(Int($numLinhas) + 15)).select
;~ 	  Sleep(600)
;~ 	  $oExcelProposta.range("F"&String(Int($numLinhas) + 3)).select
;~ 	  Sleep(900)

;~ 	  procuraClicaY("img\excelSelecionado(2).bmp",2,1,4)
;~ 	  Sleep(900)

	  $oExcelProposta.range("F2").Select
	  Sleep(1000)


;~ 	  Send("{ALTDOWN}{TAB}{ALTUP}")
;~ 	  Sleep(200)
;~ 	  $oExcelEmpresarial.range("B2").select

;~ 	  Sleep(100)
;~ 	  Send("+{END}")
;~ 	  Sleep(200)
;~ 	  Send("{DOWN}")

;~ 	  Sleep(500)
;~ 	  Send("{DOWN}")
;~ 	  Send("{HOME}")
;~ 	  Send("^v")

;~ 	  Send("{left}")

;~ 	  $oWorkbookEmpresarial.Activesheet.Columns("C:C").NumberFormat = "###"
;~ 	  Send("{CTRLDOWN}{SHIFTDOWN}{DOWN}{RIGHT}{RIGHT}{SHIFTUP}{CTRLUP}")

;~ 	  Send("^+f")
;~ 	  Sleep(500)
;~ 	  Send("p")
;~ 	  Sleep(100)
;~ 	  Send("!f")
;~ 	  Sleep(150)
;~ 	  Send("{SPACE}")
;~ 	  Sleep(100)
;~ 	  Send("^{TAB}")
;~ 	  Sleep(200)
;~ 	  Send("{ENTER}")

;~ 	  filtroDeCores()

   EndIf

EndFunc

Func filtroDeCores()
   Sleep(1000)
   $oWorkbookEmpresarial.Sheets("Plan1").Select
   Sleep(500)
   Send("^{down}")
   Sleep(800)
   Send("{LEFT 2}")
   Sleep(100)

   Send("{LALT}")
   Sleep(900)
   Send("s")
   Sleep(300)
   Send("f")
   Sleep(1600)
   $oExcelEmpresarial.Range("F1").Select
   Send("{left 3}")
   Sleep(350)

   If procuraClica("img\excelSelecionado.bmp",1) = 1 Then
	  If procuraClica("img\excelSelecionado2.bmp",1) = 1 Then
		 BlockInput(0)
		 MsgBox(0,"Erro P1","Função procuraClica não retornou valor.")
		 Exit
	  EndIf
   EndIf

   Sleep(1500)
   Send("i")
   Sleep(1500)
   If procuraClica("img\semPreenchimento.bmp",1) = 1 Then
	  BlockInput(0)
	  MsgBox(0,"Erro P5","Função procuraClica não retornou valor.")
	  Exit
   EndIf
   Sleep(800)
   $oExcelEmpresarial.Range("F1").Select
   Sleep(350)

   If procuraClica("img\excelSelecionado.bmp",1) = 1 Then
	  If procuraClica("img\excelSelecionado2.bmp",1) = 1 Then
		 BlockInput(0)
		 MsgBox(0,"Erro P2","Função procuraClica não retornou valor.")
		 Exit
	  EndIf
   EndIf

   Send("{TAB 7}")
   Sleep(200)
   Send("Vazia")
   Sleep(700)

   If _ImageSearch("img\nenhumaCorrespondencia.bmp",0,$X,$Y,0) = 1 Then
	  BlockInput(0)
	  MsgBox(0,"Atenção!","Coluna 'Nº Solic Serv' não possui valores 'Vazios'")
	  Exit
   Else
	  Send("{ENTER}")
   EndIf
   limpaClipBoard()
   Sleep(800)
   Send("{down 2}")
   Send("{CTRLDOWN}{SHIFTDOWN}{SPACE}{SPACE}{SHIFTUP}{CTRLUP}")
   Send("^c")
   Sleep(1000)
   _Excel_SheetAdd($oWorkbookEmpresarial, -1, False, 1, "Plan2")
   Send("^v")
   Sleep(600)
   $oWorkbookEmpresarial.Sheets("Plan2").Select
   $oExcelEmpresarial.Range("F2").Select
   Send("= SUBTOTAL(3;A2:A500){ENTER}")
   Sleep(400)
   $oExcelEmpresarial.Range("F2").Select
   Send("^c")
   Sleep(400)
   $numLinhas = ClipGet()
   Sleep(300)
   $oExcelEmpresarial.Range("F2").Select
   Send("{DEL}")
   Sleep(300)
   ajustaColunas($oExcelEmpresarial)
EndFunc

Func colaNaSituacao($tipo)
   Switch $tipo
	  Case 1
		 $arqu = WinActivate("Propost")
		 Sleep(800)
		 $oWorkbookProposta.Sheets("Plan2").Select
		 $oExcelProposta.Range("F"&$i).Select
		 Sleep(200)
		 Send("^v")
		 Sleep(200)

		 $oExcelProposta.Range("G"&$i).Select
		 Sleep(200)
		 Send($quantidadeItem)
		 Sleep(200)
		 Send("{ENTER}")
		 Sleep(500)

;~ 		 $return = StringRegExp($descricaoVistoria,"SIENG05S SOLICITACAO GRAVADA COM SUCESSO NUM",0,1)
;~ 		 If ($return = 1) Then ;~~Encontrou a string
		 $oExcelProposta.Range("H"&$i).Select
		 Sleep(200)
		 ClipPut($renovacao)
		 Send("^v")
		 Sleep(200)
		 Send("{ENTER}")
		 Sleep(500)
		 $renovacao = 0

		 $oExcelProposta.Range("I"&$i).Select
		 Sleep(200)
		 ClipPut($atividade)
		 Send("^v")
		 Sleep(200)
		 Send("{ENTER}")
		 Sleep(500)
		 $atividade = 0

		 $oExcelProposta.Range("J"&$i).Select
		 Sleep(200)
		 ClipPut($ocupacao)
		 Send("^v")
		 Sleep(200)
		 Send("{ENTER}")
		 Sleep(500)
		 $ocupacao = ""

		 $oExcelProposta.Range("K"&$i).Select
		 Sleep(200)
		 ClipPut($cnae)
		 Sleep(300)
		 Send("^v")
		 Sleep(200)
		 Send("{ENTER}")
		 Sleep(500)
		 $cnae = ""

		 $oExcelProposta.Range("L"&$i).Select
		 Sleep(200)
		 ClipPut($incendio)
		 Send("^v")
		 Sleep(200)
		 Send("{ENTER}")
		 Sleep(500)
		 $incendio = 0

		 $oExcelProposta.Range("M"&$i).Select
		 Sleep(200)
		 ClipPut($vendaval)
		 Send("^v")
		 Sleep(200)
		 Send("{ENTER}")
		 Sleep(500)
		 $vendaval = 0

		 $oExcelProposta.Range("N"&$i).Select
		 Sleep(200)
		 ClipPut($maiorDosRoubos)
		 Send("^v")
		 Sleep(200)
		 Send("{ENTER}")
		 Sleep(500)
		 $maiorDosRoubos = 0
;~ 		 EndIf
		 $i=$i+1

	  Case 2
		 $arqu = WinActivate("Propost")
		 Sleep(1000)
		 $oWorkbookProposta.Sheets("Plan2").Select
		 $oExcelProposta.Range("F"&$i).Select
		 Sleep(200)
		 Send("Proposta Não Cadastrada")
		 Send("{ENTER}")
		 Sleep(1000)
		 $i=$i+1
	  Case 3
		 $arqu = WinActivate("Propost")
		 Sleep(1000)
		 $oWorkbookProposta.Sheets("Plan2").Select
		 $oExcelProposta.Range("F"&$i).Select
		 Sleep(200)
		 Send("...")
		 Send("{ENTER}")
		 Sleep(1000)
		 $i=$i+1
	  Case 4
		 $arqu = WinActivate("Propost")
		 Sleep(1000)
		 $oWorkbookProposta.Sheets("Plan2").Select
		 $oExcelProposta.Range("F"&$i).Select
		 Sleep(200)
		 Send("Não cadastrada Ou Duplicada")
		 Send("{ENTER}")
		 Sleep(1000)
		 $i=$i+1
   EndSwitch
EndFunc

Func abreSiesEmissao()
   $titulo = WinGetTitle("» SIES – Sistema Especialista para Seguros - Windows Internet Explorer")

   If $titulo = "" Then
	  BlockInput(0)
	  MsgBox(0,"Erro2", "Janela do SIES não encontrada.")
	  Exit
   EndIf

   WinActivate($titulo)
   WinMove($titulo,"",0,0,@DesktopWidth,@DesktopHeight-40)
   WinSetState($titulo,"",@SW_MAXIMIZE)

   Send("{f5}")
   Sleep(3000)
   While not procuraApenas("img\btnSies.bmp");~ n achou é 0
	  ToolTip("btn sies nao encontrado",0,0)
	  Sleep(1000)
   WEnd

   ToolTip("")
   Sleep(2000)

   Send("{LCTRL}")
   Sleep(1000)
   procuraClicaX("img\linkPatrimoniais.bmp",1,1,10)
   Sleep(850)
   procuraClicaX("img\linkMrEmpresarial.bmp",1,1,10)
   Sleep(850)
   procuraClicaX("img\linkApolice.bmp",1,1,10)
   Sleep(850)
   procuraClicaX("img\linkProposta.bmp",1,1,10)
   Sleep(850)
   procuraClicaX("img\linkEmissao.bmp",2,1,10)
   Sleep(100)
    MouseMove($X+10,$Y)
    MouseClick("primary",$X+10,$Y)
   Sleep(850)
EndFunc

Func procuraClica($img, $click)
   $imgFound = _ImageSearch($img,0,$X,$Y,0)

   If $imgFound = 1 Then
	  MouseClick("primary",$X-3,$Y,$click,1);~
   Else
;~ 	  Local $cutted = StringSplit($img,'.',$STR_ENTIRESPLIT)
;~ 	  $imgFound = _ImageSearch($cutted[1]&'Low.'&$cutted[2],0,$X,$Y,0)

;~ 	  If $imgFound = 1 Then
;~ 		 MouseClick("primary",$X-3,$Y,$click,1);~
;~ 	  Else
;~ 		 return 1 ;~~ nao encontrou valor
;~ 	  EndIf

	  return 1 ;~~ nao encontrou valor
   EndIf
EndFunc

Func procuraClicaX($img, $click,$posNegati,$valor) ;~~imagem, quantos clicks, positivo ou negativo, quantidade de mudança
   $imgFound = _ImageSearch($img,0,$X,$Y,0)
   If $imgFound = 1 Then
	  Switch $posNegati
		 Case 1 ;~~positivo
			MouseClick("primary",$X+$valor,$Y,$click,1);~
		 Case 2 ;~~negativo
			MouseClick("primary",$X-$valor,$Y,$click,1);~
	  EndSwitch
   Else
;~ 	  Local $cutted = StringSplit($img,'.',$STR_ENTIRESPLIT)
;~ 	  $imgFound = _ImageSearch($cutted[1]&'Low.'&$cutted[2],0,$X,$Y,0)
;~ 	  ConsoleWrite($cutted[1]&'Low.'&$cutted[2])
;~ 	  If $imgFound = 1 Then
;~ 		 Switch $posNegati
;~ 		 Case 1 ;~~positivo
;~ 			MouseClick("primary",$X+$valor,$Y,$click,1);~
;~ 		 Case 2 ;~~negativo
;~ 			MouseClick("primary",$X-$valor,$Y,$click,1);~
;~ 	  EndSwitch
;~ 	  Else
;~ 		 return 1 ;~~ nao encontrou valor
;~ 	  EndIf

	  return 1 ;~~ nao encontrou valor
   EndIf
EndFunc

Func procuraClicaY($img, $click,$posNegati,$valor) ;~~imagem, quantos clicks, se aumenta pra mais ou menos, quantidade de mudança
   $imgFound = _ImageSearch($img,0,$X,$Y,0)
   If $imgFound = 1 Then
	  Switch $posNegati
		 Case 1 ;~~positivo
			MouseClick("primary",$X,$Y+$valor,$click,1);~
		 Case 2 ;~~negativo
			MouseClick("primary",$X,$Y-$valor,$click,1);~
	  EndSwitch
   Else
;~ 	  Local $cutted = StringSplit($img,'.',$STR_ENTIRESPLIT)
;~ 	  $imgFound = _ImageSearch($cutted[1]&'Low.'&$cutted[2],0,$X,$Y,0)
;~ 	  ConsoleWrite($cutted[1]&'Low.'&$cutted[2])
;~ 	  If $imgFound = 1 Then
;~ 		 Switch $posNegati
;~ 		 Case 1 ;~~positivo
;~ 			MouseClick("primary",$X,$Y+$valor,$click,1);~
;~ 		 Case 2 ;~~negativo
;~ 			MouseClick("primary",$X,$Y-$valor,$click,1);~
;~ 	  EndSwitch
;~ 	  Else
;~ 		 return 1 ;~~ nao encontrou valor
;~ 	  EndIf
	  return 1 ;~~ nao encontrou valor
   EndIf
EndFunc

Func procuraApenas($img) ;~~Retorno 0 nao encontrou a imagem
   $imgFound = _ImageSearch($img,0,$X,$Y,0)
   If $imgFound = 1 Then

	  return 1
   Else
;~ 	  Local $cutted = StringSplit($img,'.',$STR_ENTIRESPLIT)
;~ 	  $imgFound = _ImageSearch($cutted[1]&'Low.'&$cutted[2],0,$X,$Y,0)


;~ 	  If $imgFound = 1 Then

;~ 		 return 1
;~ 	  Else
;~ 		 return 0 ;~~ nao encontrou valor
;~ 	  EndIf
	  return 0 ;~~ nao encontrou valor
   EndIf
EndFunc

Func ajustaColunas($oExcel)
   $oExcel.Activesheet.Columns("A:A").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("B:B").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("C:C").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("D:D").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("E:E").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("F:F").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("G:G").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("H:H").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("I:I").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("J:J").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("K:K").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("L:L").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("M:M").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("N:N").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("O:O").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("P:P").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("Q:Q").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("R:R").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("S:S").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("T:T").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("U:U").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("V:V").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("W:W").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("X:X").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("Y:Y").EntireColumn.AutoFit
   $oExcel.Activesheet.Columns("Z:Z").EntireColumn.AutoFit
EndFunc

Func verificaArquivos()
   $arqu = WinActivate("Empresari")

   If $arqu = 0 Then
	  BlockInput(0)
	  MsgBox(0,"Erro 2", "Arquivo 'Empresarial' não está aberto. Por favor abra o arquivo e execute o Script novamente.")
	  Exit
   EndIf

   $arquPendente = WinActivate("Pendentes de An")

   If $arquPendente = 0 Then
	  BlockInput(0)
	  MsgBox(0,"Erro 1", "Arquivo de Propostas não está aberto. Por favor abra o arquivo e execute o Script novamente.")
	  Exit
   EndIf
EndFunc

func sair()
   Exit
EndFunc

Func limpaClipBoard()

_ClipBoard_Open(0)
_ClipBoard_Empty()
_ClipBoard_Close()
Sleep(100)
EndFunc   ;==>limpaClipBoard