#include 'protheus.ch'
#include 'parmtype.ch'
#Include 'tbiconn.ch'
#Include 'topconn.ch'
#Include 'rwmake.ch'
#Include 'Ap5Mail.ch'

//--------------------------------------------------------------------------------------
/*/{Protheus.doc} CP_EFIN1 - Rotina utilizada para o envio de email com titulos
- com vencimento em 5 dias.

@type Function
@author Michele Gois
@since 22/09/2021
@obs Cliente Anion
@version 1.0
/*/
//--------------------------------------------------------------------------------------

User Function CP_EFIN1()

	Local _nDias
	Local _cProxCli := ''
	Local _cProxLj  := ''
	Local _cEnvMail := ''
	Local _cFil     := ''
	Local _cPrefixo := ''
	Local _cTipo    := ''
	Local _cVenc    := ''
	Local _nValor   := 0
	Local _cVencReal
	Local _ntotE    := 0
	Local _cParEnv  := ''
	Local _cCarta1  := ''
	Local nI := 1

	Private  aParam := {}
	Private _oProcess	:= Nil
	Private _oHtml		:= Nil
	Private _dData 		:= Date()
	Private _dDtVenc
	Private _dDtVenc2
	Private _nTot       := 0
	Private _cNum       := ''
	Private _cParc      := ''
	Private _cCodCli    := ''
	Private _cLoja      := ''
	Private _aArray     := {}
	Private _cEmail     := ''
	Private _cNome      := ''
	Private _cTexto     := ''
	Private _nDiaSem
	Private _cUF        := ''

	For nI := 1 to  5

		If nI == 1						//ANION JANDIRA
			aAdd(aParam,{"01","01"})
		ElseIf nI == 2					//ANION MACAE
			aAdd(aParam,{"01","02"})
		ElseIf nI == 3					//MGS
			aAdd(aParam,{"02","01"})
		ElseIf nI == 4					//OFFSHORE
			aAdd(aParam,{"03","01"})
		ElseIf nI == 5					//REVESTSUL
			aAdd(aParam,{"04","01"})
		Endif

		If aParam <> Nil
			Reset Environment
			RPCSetType(3)  					//| Nao utilizar licenca

			WfPrepENV(aParam[nI][1], aParam[nI][2])
			cEmp := FWEmpName(aParam[nI][1])
			cFil := FWFilialName()

			lSchedule := .T.

			ConOut('Executando Rotina [CP_EFIN1]: EMPRESA: '+cEmp+' - '+cFil+'')

			cEmp:= ""
			cFil := ""

		EndIf

		//Tratamento para nao haver execucao concorrente
		If !(LockByName('CP_EFIN1', .T., .F.))
			ConOut("CP_EFIN1 ja esta em execucao")
			if lSchedule
				RpcClearEnv()
			Endif
			Return
		Endif

		_cParEnv  := Alltrim(GetMv("CP_ENVCART"))  //Parametro que devine se o envio da carta está Ativo/Desativado
		_cCarta1  := Substr(_cParEnv,1,1)

		If _cCarta1 == "S"   //  Se o envio estiver ativo

			_nDias   := GetMv("CP_DFINAN1")
			_dDtVenc := DaySum(_dData,_nDias) //soma _nDias a data atual
			_nDiaSem := Dow(_dData)  // recebe o dia da semana

			If _nDiaSem = 6            // 6 igual a Sexta - feira
				_dDtVenc2 := DaySum(_dData,5) //soma 5 dias a data atual
				Get2SE1()    // Select
			Else
				GetSE1()    // Select
			EndIf

			If _nTot > 0    // Quantidade de registros obtido na select

				While TRBSE1->(!EOF())

					_cFil     := TRBSE1->E1_FILIAL
					_cPrefixo := TRBSE1->E1_PREFIXO
					_cCodCli  := TRBSE1->E1_CLIENTE
					_cLoja    := TRBSE1->E1_LOJA
					_cNum     := TRBSE1->E1_NUM
					_cParc    := TRBSE1->E1_PARCELA
					_cTipo    := TRBSE1->E1_TIPO
					_cVencReal:= TRBSE1->E1_VENCREA
					_nValor   := TRBSE1->E1_SALDO
					_cVenc    := SUBSTR(_cVencReal, 7, 2) + "/" + SUBSTR(_cVencReal , 5, 2) +"/"+ SUBSTR( _cVencReal, 1, 4)
					_cEnvMail := POSICIONE ("SA1",1,xFilial("SA1")+_cCodCli+_cLoja,"A1_XENVFIN")  // Campo que determina se o cliente deve receber notificação por e-mail
					_cNome    := Alltrim(POSICIONE ("SA1",1,xFilial("SA1")+_cCodCli+_cLoja,"A1_NREDUZ"))
					_cEmail   := Alltrim(POSICIONE ("SA1",1,xFilial("SA1")+_cCodCli+_cLoja,"A1_XCOBMAI"))// E-mail de cobrança
					_cUF      := Alltrim(POSICIONE ("SA1",1,xFilial("SA1")+_cCodCli+_cLoja,"A1_EST"))

					If Empty(_cEnvMail) .or. _cEnvMail =='S'

						If Empty(_cEmail)  // Se o cliente não possui e-mail financeiro
							_cEmail   := Alltrim(POSICIONE("SA1",1,xFilial("SA1")+_cCodCli+_cLoja,"A1_XCOBMAI"))// E-mail de cobrança
						EndIf

						//Tratativa para definir %TEXTO% do WF_CP_EFIN1.htm

						If DToC(_dDtVenc) == _cVenc
							_cTexto := 'próximos cinco (5) dias'
						Else
							_cTexto := 'próximos dias'
						EndIf

						aAdd(_aArray,{_cPrefixo, _cNum, _cParc, _cTipo, _cVenc, _nValor })

						TRBSE1->(dBskip())

						_cProxCli  := TRBSE1->E1_CLIENTE    // Recebe próximo cliente da lista
						_cProxLj   := TRBSE1->E1_LOJA       // Recebe próxima loja da lista

						If _cCodCli == _cProxCli .and. _cLoja == _cProxLj  // Valida se o próximo cliente é idem ao atual

							While _cCodCli == _cProxCli .and. _cLoja == _cProxLj

								_cFil     := TRBSE1->E1_FILIAL
								_cPrefixo := TRBSE1->E1_PREFIXO
								_cNum     := TRBSE1->E1_NUM
								_cParc    := TRBSE1->E1_PARCELA
								_cTipo    := TRBSE1->E1_TIPO
								_cVencReal:= TRBSE1->E1_VENCREA
								_nValor   := TRBSE1->E1_SALDO
								_cVenc    := SUBSTR(_cVencReal, 7, 2) + "/" + SUBSTR(_cVencReal , 5, 2) +"/"+ SUBSTR( _cVencReal, 1, 4)

								aAdd(_aArray,{_cPrefixo, _cNum, _cParc, _cTipo, _cVenc, _nValor })

								TRBSE1->(dBskip())

								_cProxCli  := TRBSE1->E1_CLIENTE  // Recebe próximo cliente da lista
								_cProxLj   := TRBSE1->E1_LOJA     // Recebe próxima loja da lista
							EndDo

						EndIf
						xEnvEmail()
						_ntotE ++		//Contador para quantidade de emails enviados
					Else
						TRBSE1->(dBskip())
					EndIf
					_aArray := {}
				EndDo

			EndIf

			If lSchedule
				RpcClearEnv()
			Endif			

		EndIf
		
		ConOut('Finalizando rotina [CP_EFIN1]: EMPRESA: '+cEmp+' - '+cFil+'')

	Next nI
	
Return

//-----------------------------------------------------------------------------------------------------------------------//
Static Function xEnvEmail()

	Local cServer		:= ALLTRIM(GetMv('MV_RELSERV',,''))	//|	Nome do servidor de envio de e-mail
	Local cAccount		:= ALLTRIM(GetMv('MV_RELACNT',,''))	//|	Conta a ser utilizada no envio de e-mail
	Local cPassword		:= ALLTRIM(GetMv('MV_RELPSW',,''))	//|	Senha da conta de e-mail para envio
	Local lAutentica	:= GetMv('MV_RELAUTH',,'')	//|	Determina se o Servidor de Email necessita de Autenticacao
	Local _lOk			:= .T.
	Local _lenvia		:= .T.
	Local _cEnv 		:= AllTrim(Upper(GetEnvServer()))
	Local _lColor		:= .F.
	Local _cBody		:= ''
	Local _cCco     	:= Alltrim(GetMv("CP_CFINANC"))
	Local a				:= 0

	_cCabec := '<html>'
	_cCabec += '<body>'

	_oProcess:= TWFProcess():New('','Titulos Provisorios')
	_oProcess:NewTask('Inicio','/workflow/Html/emailfin/WF_CP_EFIN1.htm')

	For a:=1 to Len(_aArray)

		_cCabec += '<tr>'

		_cCabec += '<td align=center style="border-top:solid #dddddd 1.0pt;border-left:solid #d5d5d5 1.0pt;border-bottom:none;border-right:solid #f2f2f2 1.0pt' + IIf(_lColor,';background:whitesmoke','') + '">'
		_cCabec += AllTrim(_aArray[a][1])
		_cCabec += '</td>'

		_cCabec += '<td align=center style="border-top:solid #dddddd 1.0pt;border-left:solid #d5d5d5 1.0pt;border-bottom:none;border-right:solid #f2f2f2 1.0pt' + IIf(_lColor,';background:whitesmoke','') + '">'
		_cCabec += AllTrim(_aArray[a][2])
		_cCabec += '</td>'

		_cCabec += '<td align=center style="border-top:solid #dddddd 1.0pt;border-left:solid #d5d5d5 1.0pt;border-bottom:none;border-right:solid #f2f2f2 1.0pt' + IIf(_lColor,';background:whitesmoke','') + '">'
		_cCabec += AllTrim(_aArray[a][3])
		_cCabec += '</td>'

		_cCabec += '<td align=center style="border-top:solid #dddddd 1.0pt;border-left:solid #d5d5d5 1.0pt;border-bottom:none;border-right:solid #f2f2f2 1.0pt' + IIf(_lColor,';background:whitesmoke','') + '">'
		_cCabec += AllTrim(_aArray[a][4])
		_cCabec += '</td>'

		_cCabec += '<td align=center style="border-top:solid #dddddd 1.0pt;border-left:solid #d5d5d5 1.0pt;border-bottom:none;border-right:solid #f2f2f2 1.0pt' + IIf(_lColor,';background:whitesmoke','') + '">'
		_cCabec += AllTrim(_aArray[a][5])
		_cCabec += '</td>'

		_cCabec += '<td align=right style="border-top:solid #dddddd 1.0pt;border-left:solid #d5d5d5 1.0pt;border-bottom:none;border-right:solid #f2f2f2 1.0pt' + IIf(_lColor,';background:whitesmoke','') + '">'
		_cCabec += Transform(_aArray[a][6], '@E 999,999.99' )
		_cCabec += '</td>'
		_cCabec += '</tr>' + CRLF

		_lColor := !_lColor

	Next a

	GeraHtml()

	_cBody 	:= _oHtml:HtmlCode()
	_oProcess:Free()
	_oProcess	:= Nil

	_cTitulo := 'Comunicado sobre TÍTULOS A VENCER'
	_cRod := '<br><br><br>'
	_cRod += '</body>'
	_cRod += '</html>'

	If _lenvia

		//If alltrim(_cEnv) $("PRODUCTION|WFFINANCEIRO|WS")
		If alltrim(_cEnv) $("WFFINANCEIRO")
			cDe     := cAccount
			cPara   := _cEmail
			cCc     :=  ""
			cCo     := _cCco
			cAssunto:= _cTitulo
			cAnexo  := ""
		Else
			cDe     := cAccount
			cPara   := 'michele.gois@cyberpolos.com.br'
			cCc     :=  ""
			cCo     := ""//_cCco
			cAssunto:= 'HOMOLOGAÇÃO - '+ _cTitulo
			cAnexo  := ""
		Endif

		CONNECT SMTP SERVER cServer ACCOUNT cAccount PASSWORD cPassword RESULT lOk

		//Autentica, se necessário
		If lAutentica
			If !MailAuth(cAccount,cPassword)
				ConOut("Erro de autenticacao no envio de email. Processo CP_EFIN1")
				DISCONNECT SMTP SERVER
				Return .F.
			EndIf
		EndIf

		_lOk := MailSend (cDe, {cPara}, {cCc}, {cCo}, cAssunto, _cBody, {''}, .T. )

		If !_lOk
			GET MAIL ERROR _cErro
			Conout("Erro ao enviar email para cliente")
			ConOut(_cErro)
			DISCONNECT SMTP SERVER

			EnvMailErr(cServer,cAccount,cPassword,lAutentica,cDe,cAssunto,_cBody)
			Return .F.
		EndIf

		DISCONNECT SMTP SERVER

	Endif

Return

//-----------------------------------------------------------------------------------------------------------------------//
Static Function GeraHtml()

	Local _cDia     := ''
	Local _cMes		:= ''
	Local _cAno		:= ''
	Local _cCont1   := ''
	Local _cNomeEmp	:= ''

	_cCont1   := Alltrim(GetMv("CP_CFINAN1"))

	_cDia     := Day(_dData)
	_cMes	  := MesExtenso(_dData)
	_cAno     := Year(_dData)
	_cNomeEmp := SM0->M0_NOMECOM

	_oHtml   := _oProcess:oHtml
	_oHtml:Valbyname('RODAPE' 		, Alltrim(_cNomeEmp))
	_oHtml:Valbyname('LISTAITENS' 	, _cCabec)
	_oHtml:Valbyname('CLIENTE' 		, _cNome)
	_oHtml:Valbyname('DIA' 			, _cDia)
	_oHtml:Valbyname('MES' 			, _cMes)
	_oHtml:Valbyname('ANO' 			, _cAno)
	_oHtml:Valbyname('CODIGO' 		, _cCodCli)
	_oHtml:Valbyname('UF' 			, _cUF)
	_oHtml:Valbyname('CONTATOS1' 	, _cCont1)

Return

//-----------------------------------------------------------------------------------------------------------------------//
Static Function GetSE1()

	Local _lRet	:= .T.

	If Select('TRBSE1') > 0
		TRBSE1->(DbCloseArea())
	EndIf

	BeginSql Alias 'TRBSE1'
		SELECT
			E1_FILIAL,
			E1_NUM,
			E1_PARCELA,
			E1_CLIENTE,
			E1_LOJA,
			SE1.R_E_C_N_O_ AS SE1_RECNO,
			E1_VENCREA,
			E1_SALDO,
			E1_TIPO,
			E1_PREFIXO,
			E1_VALOR,
			E1_NOMCLI
		FROM
			%Table:SE1% SE1
		WHERE
			E1_FILIAL = %xFilial:SE1%
			AND E1_VENCREA = %Exp:_dDtVenc%
			AND LTRIM(RTRIM(E1_TIPO)) not in ('NCC', 'RA', 'TCC')
			AND E1_SITUACA not in ('2', '7')
			AND E1_SALDO > 0
			AND SE1.%NotDel%
		ORDER BY
			E1_CLIENTE,
			E1_LOJA,
			E1_VENCREA,
			E1_NUM,
			E1_PARCELA
	EndSQL

	If TRBSE1->(EOF())
		_lRet := .F.
		ConOut('CP_EFIN1 Não foram encontrados registros para os parâmetros informados.')
	Else
		Count to _nTot
		TRBSE1->(DbGoTop())
	EndIf

Return _lRet

//-----------------------------------------------------------------------------------------------------------------------//

Static Function Get2SE1()

	Local _lRet	:= .T.

	If Select('TRBSE1') > 0
		TRBSE1->(DbCloseArea())
	EndIf

	BeginSql Alias 'TRBSE1'
		SELECT
			E1_FILIAL,
			E1_NUM,
			E1_PARCELA,
			E1_CLIENTE,
			E1_LOJA,
			SE1.R_E_C_N_O_ AS SE1_RECNO,
			E1_VENCREA,
			E1_SALDO,
			E1_TIPO,
			E1_PREFIXO,
			E1_VALOR,
			E1_NOMCLI
		FROM
			%Table:SE1% SE1
		WHERE
			E1_FILIAL = %xFilial:SE1%
			AND E1_VENCREA BETWEEN %Exp:_dDtVenc% AND %Exp:_dDtVenc2%
			AND LTRIM(RTRIM(E1_TIPO)) not in ('NCC', 'RA', 'TCC')
			AND E1_SITUACA not in ('2', '7')
			AND E1_SALDO > 0
			AND SE1.%NotDel%
		ORDER BY
			E1_CLIENTE,
			E1_LOJA,
			E1_VENCREA,
			E1_NUM,
			E1_PARCELA
	EndSQL

	If TRBSE1->(EOF())
		_lRet := .F.
		ConOut('CP_EFIN1 Não foram encontrados registros para os parâmetros informados.')
	Else
		Count to _nTot
		TRBSE1->(DbGoTop())
	EndIf

Return _lRet

//-----------------------------------------------------------------------------------------------------------------------//
Static Function EnvMailErr(cServer,cAccount,cPassword,lAuth,cDe,cAssunto,_cBody)
	Local _lOk			:= .T.
	Local _cParaErr 	:= 'michele.gois@cyberpolos.com.br'
	Local _cCc			:= ''
	Local _cCco		:= ''

	Private _cErro	:= ''

	CONNECT SMTP SERVER cServer ACCOUNT cAccount PASSWORD cPassword RESULT _lOk

	If lAuth
		If !MailAuth(cAccount,cPassword)
			ConOut("Erro de autenticacao no envio de email. Processo CP_EFIN1")
			DISCONNECT SMTP SERVER
			Return
		EndIf
	EndIf

	_lOk := MailSend ( cDe, {_cParaErr}, {_cCc}, {_cCco}, 'Erro ao enviar ' + cAssunto, _cBody, {''}, .T. )

	If !_lOk

		GET MAIL ERROR _cErro
		Conout("Erro ao enviar email de ERRO DA ROTINA CP_EFIN1")
		ConOut(_cErro)
	EndIf

	DISCONNECT SMTP SERVER
Return

