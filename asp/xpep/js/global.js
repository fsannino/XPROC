function mOvr(src,clrOver) 
{
	if (!src.contains(event.fromElement)) {
		src.style.cursor = 'hand';
		src.bgColor = clrOver;
	}
}
function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
		src.style.cursor = 'default';
		src.bgColor = clrIn;
	}
}
function mClk(src) {
	if(event.srcElement.tagName=='TD'){
		src.children.tags('A')[0].click();
	}
}

var blnData;
function validaData(strData, strCampo, strNomeCampo)
{		
	var strDTPadrao = new String;				
	var strDTInicio = new String;
	var strDTFinal  = new String;
	var strDataSel = strData;								
	var strDia;
	var strMes;
	var strAno;		
	var strDiaIni;
	var strMesIni;
	var strAnoIni;		
	var strDiaFim;
	var strMesFim;
	var strAnoFim;		
	blnData = false;								
			
	strDTPadrao = strData.split('/')				
	strDia = strDTPadrao[0] 
	strMes = strDTPadrao[1]
	strAno = strDTPadrao[2]
	
	strDTInicio = document.forms[0].pDtInicioAtiv.value;		
	strDTInicio = strDTInicio.split('/')				
	strDiaIni = strDTInicio[0] 
	strMesIni = strDTInicio[1]
	strAnoIni = strDTInicio[2]
	
	strDTFinal = document.forms[0].pDtFimAtiv.value;
	strDTFinal = strDTFinal.split('/')				
	strDiaFim = strDTFinal[0] 
	strMesFim = strDTFinal[1]
	strAnoFim = strDTFinal[2]
				
	if ((strCampo == 'txtDtLimiteAprov') || (strCampo == 'txtDTAprovacao_PAC'))
	{
		if ((strAno > strAnoIni)) 
		{
			alert('A data ' + strDataSel + ' preenchida no campo "' + strNomeCampo + '", deve ser menor do que a Data de Início do Cronograma!');
			blnData = true; 
			return(blnData);
		}
		else 
		{
			if ((strMes > strMesIni) && (strAno >= strAnoIni))
			{
				alert('A data ' + strDataSel + ' preenchida no campo "' + strNomeCampo + '", deve ser menor do que a Data de Início do Cronograma!');
				blnData = true; 
				return(blnData);
			}
			else 
			{
				if ((strDia >= strDiaIni) && (strMes >= strMesIni) && (strAno >= strAnoIni))
				{
					alert('A data ' + strDataSel + ' preenchida no campo "' + strNomeCampo + '", deve ser menor do que a Data de Início do Cronograma!');
					blnData = true; 
					return(blnData);
				}
			}
		}
	}
	else
	{
		if ((strAno < strAnoIni)||(strAno > strAnoFim)) 
		{
			alert('A data ' + strDataSel + ' preenchida no campo "' + strNomeCampo + '", deve estar dentro do intervalo Data Inicio/Data de Término!');
			blnData = true; 
			return(blnData);
		}
		else 
		{
			if (((strMes < strMesIni) && (strAno <= strAnoIni))||((strMes > strMesFim) && (strAno >= strAnoFim)))
			{
				alert('A data ' + strDataSel + ' preenchida no campo "' + strNomeCampo + '", deve estar dentro do intervalo Data Inicio/Data de Término!');
				blnData = true; 
				return(blnData);
			}
			else 
			{
				if (((strDia < strDiaIni) && (strMes <= strMesIni) && (strAno <= strAnoIni))||((strDia > strDiaFim) && (strMes >= strMesFim) && (strAno >= strAnoFim))) 
				{
					alert('A data ' + strDataSel + ' preenchida no campo "' + strNomeCampo + '", deve estar dentro do intervalo Data Inicio/Data de Término!');
					blnData = true; 
					return(blnData);
				}
			}
		}
	}
}	