//Default browsercheck, added to all scripts!
function checkBrowser(){
	this.ver=navigator.appVersion
	this.dom=document.getElementById?1:0
//	this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom)?1:0; é o que tinha antes
//  alterei a linha abaixo para checagem tb do IE 6 Beta :-) --> 
	this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom || this.ver.indexOf("MSIE 6.0b")>-1 || this.ver.indexOf("MSIE 6")>-1 && this.dom)?1:0;
	this.ie4=(document.all && !this.dom)?1:0;
	this.ns5=(this.dom && parseInt(this.ver) >= 5) ?1:0;
	this.ns4=(document.layers && !this.dom)?1:0;
	this.bw=(this.ie5 || this.ie4 || this.ns4 || this.ns5)
	return this
}
var bw=new checkBrowser()
var qtd=0
var raiz=''
var posicao
//raiz = 'xproc'
raiz = 'http://s6000ws10.corp.petrobras.biz/xproc'

//Ie var
var explorerev=''

var lvl=''
var offline= '';
var online= 'http://www.aforbes.com.br';
var cnt=0;
var off_cnt=0

function getLevel(){
	var cnt=0;
	var addr='';
	var tmp= location.href;
	if(tmp.indexOf('file:')>-1 || tmp.charAt(1)==':') addr= offline;
	else if(tmp.indexOf('http:')>-1) addr= online;
	for(var i=0;i<addr.length;i++){
		if(addr.charAt(i)=='\/'){
			off_cnt+=1
		}
	}
	for(var i=0;i<tmp.length;i++){
		if(tmp.charAt(i)=='\/'){
			cnt+=1;
			if(cnt>off_cnt)lvl+='../'
		}
	}
}
getLevel();

function goMenus(){
/********************************************************************************
Variables to set.

Se lembre isso para fixar para fontsize e para fonttype o jogo isso no stylesheet  
sobre!
********************************************************************************/

//Fazendo um objeto do menu
oMenu=new menuObj('oMenu') //Coloque um nome para o menu. Deve ser unico para cada menu
//Configuração das variáveis do objeto menu

//Style variables NOTE: O stylesheet foram afastados. Use isto ao invés! (alguns estilos estão lá através de falta, como position:absolute ++)
oMenu.clMain='padding:4px; font-family:arial; font-size:11px; font-weight:bold; text-align:center; border: solid; border-width: 0pt 0pt 0pt 1pt; border-color: #FFFFFF #FFFFFF #FFFFFF #DFDBBF' //Style do menu principal (menu título)
oMenu.clSub='padding:5px; font-family:arial; font-size:11px; font-weight:normal; text-indent: 2pt; border-color: #FFFFFF #FFFFFF #FFFFFF; border-style: solid; border-top-width: 0px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px' //Style dos submenus
oMenu.clSubSub='padding:4px; font-family:arial; font-size:11px; font-weight:normal; text-indent: 2pt; border-color: #FFFFFF #cccccc #cccccc; border-style: solid; border-top-width: 0px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px' //Style dos subsubmenus
oMenu.clAMain='text-decoration:none; color:#ffffff' //Style da fonte do menu principal (menu título)
oMenu.clASub='text-decoration:none; color:#ffffff' //Style da fonte do submenu
oMenu.clASubSub='text-decoration:none; color:#ffffff' //Style da fonte do subsubmenu

//Background bar properties
oMenu.backgroundbar=0 //Set o 0 para nenhum backgrondbar
oMenu.backgroundbarfromleft=0 //A colocação esquerda do backgroundbar em pixel ou%
oMenu.backgroundbarfromtop=0 //A colocação de topo do backgroundbar em pixel ou%
oMenu.backgroundbarsize=0 //O tamanho da barra em pixel ou%
oMenu.backgroundbarcolor=0 //O backgroundcolor da barra

//Tamanho dos menus principais
oMenu.mainheight=22 //A altura do menuitems principal em pixel ou%
oMenu.mainwidth=80 //A largura do menuitems principal em pixel ou%

/*Estas são variáveis novas. Neste exemplo eles são fixos como a versão prévia*/
//oMenu.subwidth=oMenu.mainwidth // ** NEW ** A largura dos submenu (largura igual a do menu)
//oMenu.subheight=oMenu.mainheight // ** NEW ** A altura dos submenu (largura igual a do menu)
oMenu.subheight=25 //Caso você não queira a altura igual a do menu
oMenu.subwidth=120 //Caso você não queira a largura igual a do menu


//oMenu.subsubwidth=oMenu.mainwidth // ** NEW ** A largura do subsubmenus em pixel ou% 
//oMenu.subsubheight=oMenu.subheight //** NEW ** A altura se o subsubitems em pixel ou% 
oMenu.subsubheight=25 //Caso você não queira a altura igual a do submenu
oMenu.subsubwidth=150 //Caso você não queira a largura igual a do submenu

//Escrevendo fora o estilo para o menu (deixe esta linha!)
oMenu.makeStyle()

oMenu.subplacement=oMenu.mainheight //** NEW ** altura dos menus que irão aparecer
//oMenu.subsubXplacement=oMenu.subwidth/2 //** NEW ** distância dos submenus em relação aos links baseada nos submenu
oMenu.subsubYplacement=5 //** NEW ** Altura dos subsubmenus em relação ao links do submenu
oMenu.subsubXplacement=120 //** NEW ** distância dos submenus em relação aos links

//Cores dos Background dos menus
oMenu.mainbgcoloroff='#0066FF' //Background do menu principal
oMenu.mainbgcoloron='#0099FF' //Background do mouseover do menu principal
oMenu.subbgcoloroff='#003399' //Background do submenu
oMenu.subbgcoloron='#0066CC' //Background do mouseover do submenu
oMenu.subsubbgcoloroff='#003399' //Background do subsubmenu
oMenu.subsubbgcoloron='#0066CC' //Background do mouseover do subsubmenu
oMenu.stayoncolor=0 //Você quer os menus para ficar no mouseovered colora quando clicado? 
                    //Você quer os menus para ficar no mouseovered colora quando clicou?

//Velocidade dos menus
oMenu.menuspeed=35 //Velocidade como os menus descem 
oMenu.menusubspeed=30 //Velocidade como os submenus descem

oMenu.menurows=1 //Fixe a 0 se você quiser filas e para 1 se você quer colunas

oMenu.menueventon="mouse" //"mouse" para fazer o efeito MouseOver, "clicar" para fazer o efeito MouseOnClick
oMenu.menueventoff="mouse" //"mouse" para fazer o efeito MouseOut, "clicar" para fazer o efeito MouseOnClick

//Posição dos menus principais

//Example in %:
//oMenu.menuplacement=new Array("20%","40%","60%","50%","65%") //Se lembre de fazer as ordens conter tantos valores quanto você tem menuitems principal

//Example in px: (remember to use the ' ' around the numbers)
//oMenu.menuplacement=new Array(10,200,300,400,500)

//Example right beside eachother (only adding the pxbetween variable)
//oMenu.menuplacement=new Array('176','280','380','480','560','641') //configurar o px aqui!!!!!!!
// cada valor e o valor de cada opção do menu principal -- 
// a diferença entre eles deverá ser de 80/120 conforme o param oMenu.mainwidth=120 - uma linha ??
//oMenu.menuplacement=new Array('0','80','160','240','320','400','480','560') //os primeiros 3 números são a posição horizontal dos menus
oMenu.menuplacement=new Array('0','80') //os primeiros 3 números são a posição horizontal dos menus


//Se você usa o "direito ao lado de eachother" você hipocrisia que quanto pixel deveriam estar entre cada aqui
oMenu.pxbetween=0 //in pixel or %

//E você pode fixar onde deveria começar da esquerda aqui
oMenu.fromleft=45 //in pixel or %

//Altura do menu em relação ao topo do browser
oMenu.fromtop=80 //in pixel or %

/********************************************************************************
Construindo os menus
********************************************************************************/

//MAIN 0 - Empresa

//Main items:
// makeMain(MAIN_NUM,'TEXT','LINK','FRAME_TARGET') (set link to 0 if you want submenus of this menu item)
//MAIN 2 - Seguros

posicao=0;
oMenu.makeMain(posicao,'PERFIL',0)

	qtd_3=2;
	oMenu.makeSub(posicao,0,'Criação R/3 Macro',0,qtd_3)
	oMenu.makeSub(posicao,1,'Criação R/3 Micro',0,qtd_3)

	qtd_33=1
	oMenu.makeSubSub(posicao,0,0,'Em Criação->Criado R/3',raiz + '/asp/macroperfil/selec_valida_status5.asp',qtd_33)
	
	qtd_33=1
	oMenu.makeSubSub(posicao,1,0,'Em Criação->Criado R/3',raiz + '/asp/microperfil/selec_valida_micro2.asp',qtd_33)


posicao=1;
oMenu.makeMain(posicao,'CONSULTA',0)

    qtd_6=8;
	oMenu.makeSub(posicao,0,'Processo',0,qtd_6)
	oMenu.makeSub(posicao,1,'Cenário',0,qtd_6)
	oMenu.makeSub(posicao,2,'Função',0,qtd_6)	
	oMenu.makeSub(posicao,3,'Perfil',0,qtd_6)
	oMenu.makeSub(posicao,4,'Curso',0,qtd_6)
	oMenu.makeSub(posicao,5,'Usuário',raiz + '/asp/consulta_usuario.asp',qtd_6)
	oMenu.makeSub(posicao,6,'Trans_Duplicada',0,qtd_6)
	oMenu.makeSub(posicao,7,'Escopo',0,qtd_6)	
		
	pos_Sub = 0
	qtd_60=13
	oMenu.makeSubSub(posicao,pos_Sub,0,'Mega-Processo',raiz + '/asp/consulta_mega_processo.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Processo',raiz + '/asp/consulta_processo.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Sub-Processo',raiz + '/asp/consulta_sub.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,3,'Agrup.(Mstr List R3)',raiz + '/asp/consulta_modulo.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,4,'Atividade',raiz + '/asp/consulta_atividade.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,5,'Transações',raiz + '/asp/consulta_trans.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,6,'Empresa',raiz + '/asp/consulta_empresa.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,7,'Escopo',raiz + '/asp/rel_agrativtran.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,8,'Decomposição - Modelo 1',raiz + '/asp/rel_geral.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,9,'Decomposição - Modelo 2',raiz + '/asp/consulta.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,10,'Decomposição - Modelo 3',raiz + '/asp/rel_geral2.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,11,'Decomposição Empresa',raiz + '/asp/rel_megaemp.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,12,'Erros de Importação',raiz + '/asp/rel_bug.asp',qtd_60)
	
	pos_Sub = 1
	qtd_61=13
	oMenu.makeSubSub(posicao,pos_Sub,0,'Cenário',raiz + '/asp/cenario/rel_geral.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Cenario e Trans', raiz + '/asp/cenario/rel_cond.asp',qtd_61)		
	oMenu.makeSubSub(posicao,pos_Sub,2,'Classe',raiz + '/asp/cenario/consulta_classe.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,3,'Status',raiz + '/asp/cenario/rel_status.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,4,'Status-Excel',raiz + '/asp/cenario/gera_rel_status_excel_total.asp',qtd_61,"BLANK")
	oMenu.makeSubSub(posicao,pos_Sub,5,'Status por Período',raiz + '/asp/cenario/rel_status_periodo.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,6,'Cenário X Cenário',raiz + '/asp/cenario/rel_relac_cenario.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,7,'Cenário Ent-Saída',raiz + '/asp/cenario/rel_entsai.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,8,'Problemas Status',raiz + '/asp/cenario/rel_confcenario.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,9,'Rel Cenario',raiz + '/asp/cenario/rel_cenario.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,10,'Cen c/ Somatorio',raiz + '/asp/cenario/rel_cenario_quebra.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,11,'Cenarios REFAP',raiz + '/asp/cenario/rel_geral_refap.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,12,'Rel Cenarios sem Assunto', raiz + '/asp/cenario/sel_cenario_altera_assunto.asp?txtOPT=2',qtd_61)	
	
	pos_Sub = 2
	qtd_62=10
	oMenu.makeSubSub(posicao,pos_Sub,0,'Fun-Mega-Trans',raiz + '/asp/funcao/seleciona_Mega3.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Func-Trans',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=5',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Relatório Função',raiz + '/asp/funcao/rel_geral_funcao.asp?pOPT=1',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,3,'Função (Coluna)',raiz + '/asp/funcao/rel_geral_funcao.asp?pOPT=2',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,4,'Funções Conflitantes',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=8',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,5,'Rel Transação X Função',raiz + '/asp/funcao/consulta_func_trans.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,6,'Rel Funçao x Transação',raiz + '/asp/funcao/rel_func_trans_sem_rep.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,7,'Rel Função sem Assunto',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=7',qtd_62)	
	oMenu.makeSubSub(posicao,pos_Sub,8,'Orientações Geral',raiz + '/asp/orie_mape/relat_ori_gerais_mapeamento.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,9,'Orientações Mega',raiz + '/asp/orie_mape/seleciona_funcao.asp?pOpt=RM',qtd_62)

	pos_Sub = 3
    qtd_62=4
	oMenu.makeSubSub(posicao,pos_Sub,0,'Macro',raiz + '/asp/macroperfil/rel_geral_macro.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Solicitação Micro',raiz + '/asp/mIcroperfil/rel_mIcro.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Micro_R3',raiz + '/asp/mIcroperfil/rel_mIcro_r3.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,3,'Consulta Lote', raiz + '/asp/golive/consulta_lote.asp?pAcao=C',qtd_62)		

	pos_Sub = 4
	qtd_62=5
	oMenu.makeSubSub(posicao,pos_Sub,0,'Curso',raiz + '/asp/curso/relat_geral_curso.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Curso x Função',raiz + '/asp/curso/seleciona_curso_rel.asp?option=2',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Curso x Transação',raiz + '/asp/curso/seleciona_curso_rel.asp?option=1',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,3,'Pré Requisito',raiz + '/asp/curso/seleciona_curso_rel.asp?option=4',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,4,'Rel.Compl.Cursos',raiz + '/asp/curso/curso_prerequisito.asp',qtd_62)
	
	pos_Sub = 6
    qtd_62=3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Transação com Dono(s)',raiz + '/asp/exibe_dono.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Transação sem Dono',raiz + '/asp/exibe_sem_dono.asp',qtd_62)
	//oMenu.makeSubSub(posicao,pos_Sub,2,'Transação Dupl old',raiz + '/asp/consulta_transacao_outro_mega_old.asp',qtd_62)
	//oMenu.makeSubSub(posicao,pos_Sub,3,'Transação Dupl old2',raiz + '/asp/consulta_transacao_outro_mega_old2.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Transação X Mega',raiz + '/asp/selec_rel_transmega.asp',qtd_62)

	pos_Sub = 7
    qtd_62=4
	oMenu.makeSubSub(posicao,pos_Sub,0,'Cenário: em uma data', raiz + '/asp/cenario/consulta_escopo.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Cenário: entre datas', raiz + '/asp/cenario/consulta_escopo_entre_datas.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Transação: em uma data', raiz + '/asp/consulta_transacao_decomp.asp?tipo=1',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,3,'Transação: entre datas', raiz + '/asp/consulta_transacao_decomp.asp?tipo=2',qtd_62)
/********************************************************************************
End menu construction
********************************************************************************/
		
		
//When all the menus are written out we initiates the menu
oMenu.construct()
}

/********************************************************************************
Object constructor and object functions
********************************************************************************/
function makePageCoords(){
	this.x=0;this.x2=(bw.ns4 || bw.ns5)?innerWidth:document.body.offsetWidth-20;
	this.y=0;this.y2=(bw.ns4 || bw.ns5)?innerHeight:document.body.offsetHeight-5;
	this.x50=this.x2/2;	this.y50=this.y2/2;
	return this;
}
function makeMenu(parent,obj,nest,type,num,subnum,subsubnum){
    nest=(!nest) ? '':'document.'+nest+'.'
   	this.css=bw.dom? document.getElementById(obj).style:bw.ie4?document.all[obj].style:bw.ns4?eval(nest+"document.layers." +obj):0;					
	this.evnt=bw.dom? document.getElementById(obj):bw.ie4?document.all[obj]:bw.ns4?eval(nest+"document.layers." +obj):0;		
	this.height=bw.ns4?this.css.document.height:this.evnt.offsetHeight
	this.width=bw.ns4?this.css.document.width:this.evnt.offsetWidth
	this.moveIt=b_moveIt; this.bgChange=b_bgChange;	
	this.clipTo=b_clipTo;
	this.parent=parent;
	this.active=0;
	this.nssubover=0
	if(type==0){
		this.evnt.onmouseover=new Function("mmover("+num+","+this.parent.name+")");
		this.evnt.onmouseout=new Function("mmout("+num+","+this.parent.name+")");
	}else if(type==1){
		this.clipIn=b_clipIn;
		this.clipOut=b_clipOut;
		this.clipy=0
		if(bw.ns4 && this.parent.menueventoff=="mouse"){
			this.evnt.onmouseout=new Function("setTimeout('if(!"+this.parent.name+"["+num+"].nssubover)"+this.parent.name+".hideactive("+num+");',100)")
			this.evnt.onmouseover=new Function(this.parent.name+"["+num+"].nssubover=true")
		}
	}else if(type==2){
		this.evnt.onmouseover=new Function("submmover("+num+","+subnum+","+this.parent.name+")");
		this.evnt.onmouseout=new Function("submmout("+num+","+subnum+","+this.parent.name+")");
	}else if(type==3){
		this.evnt.onmouseover=new Function("subsubmmover("+num+","+subnum+","+subsubnum+","+this.parent.name+")");
		this.evnt.onmouseout=new Function("subsubmmout("+num+","+subnum+","+subsubnum+","+this.parent.name+")");
	}
	this.tim=100
    this.obj = obj + "Object"; 	eval(this.obj + "=this")	
	return this
}
function b_clipTo(t,r,b,l,h){if(bw.ns4){this.css.clip.top=t;this.css.clip.right=r
this.css.clip.bottom=b;this.css.clip.left=l; this.clipx=r;
}else{this.css.clip="rect("+t+","+r+","+b+","+l+")"; this.clipx=r;;
if(h){ if(bw.ie4 || bw.ie5){ this.css.height=b; this.css.width=r}}}}
function b_moveIt(x,y){this.x=x; this.y=y; this.css.left=this.x;this.css.top=this.y}
function b_bgChange(color){if(bw.dom || bw.ie4) this.css.backgroundColor=color;
else if(bw.ns4) this.css.bgColor=color}
function b_clipIn(speed){
	if(this.clipy>0){
		this.clipy-=speed
		if(this.clipy<0) this.clipy=0
		this.clipTo(0,this.clipx,this.clipy,0,1)
		this.tim=setTimeout(this.obj+".clipIn("+speed+")",10)
	}else{this.clipy=0; this.clipTo(0,this.clipx,this.clipy,0,1)}	
}
function b_clipOut(speed){
	if(this.clipy<this.clipheight){
		this.clipy+=speed
		this.clipTo(0,this.clipx,this.clipy,0,1)
		this.tim=setTimeout(this.obj+".clipOut("+speed+")",10)
	}else{this.clipy=this.clipheight; this.clipTo(0,this.clipx,this.clipy,0,1)}
}
//Page variable, holds the width and height of the document. (see documentsize tutorial on bratta.com/dhtml)
var page=new makePageCoords()

/********************************************************************************
Checking if the values are % or not.
********************************************************************************/
function checkp(num,lefttop){
	if(num){
		if(num.toString().indexOf("%")!=-1){
			if(this.menurows)num=(page.x2*parseFloat(num)/100)
			else num=(page.y2*parseFloat(num)/100)
		}else num=parseFloat(num)
	}else num=0
	return num
}
/********************************************************************************
Menu object, constructing menu ++
********************************************************************************/
function menuObj(name){
	this.makeStyle=makeStyle;
	this.makeMain=makeMain;
	this.makeSub=makeSub;
	this.makeSubSub=makeSubSub
	this.mainmenus=0; 
	this.submenus=new Array()
	this.construct=constructMenu;
	this.checkp=checkp;
	this.name=name;
	this.menumain=menumain;
	this.hidemain=hidemain;
	this.hideactive=hideactive;
	this.menusub=menusub;
	this.hidesubs=hidesubs;
}
function constructMenu(){
	bw=new checkBrowser()
	page=new makePageCoords()
	//Checking numbers for %
	this.mainwidth=checkp(this.mainwidth,0)
	this.mainheight=checkp(this.mainheight,1)
	this.subplacement=checkp(this.subplacement,1)
	this.subwidth=checkp(this.subwidth,0)
	this.subheight=checkp(this.subheight,1)
	this.subsubwidth=checkp(this.subsubwidth,0)
	this.subsubheight=checkp(this.subsubheight,1)
	this.subsubXplacement=checkp(this.subsubXplacement,1)
	this.subsubYplacement=checkp(this.subsubYplacement,1)
	if(this.backgroundbar){ //Backgroundbar part
		this.oBackgroundbar=new makeMenu(this,'div'+this.name+'Backgroundbar','',-1)
		this.oBackgroundbar.moveIt(this.checkp(this.backgroundbarfromleft,0),this.checkp(this.backgroundbarfromtop,1))
		if(this.menurows) this.oBackgroundbar.clipTo(0,this.checkp(this.backgroundbarsize),this.mainheight,0,1)
		else this.oBackgroundbar.clipTo(0,this.mainwidth,this.checkp(this.backgroundbarsize),0,1)
		this.oBackgroundbar.bgChange(this.backgroundbarcolor)
	}
	this.x=this.checkp(this.fromleft,0); this.y=this.checkp(this.fromtop,1);
	for(i=0;i<this.mainmenus;i++){
		this[i]=new makeMenu(this,'div'+this.name+'Main'+i,'',0,i)
		this[i].clipTo(0,this.mainwidth,this.mainheight,0,1)
		if(this.menuplacement!=0){
			if(this.menurows) this.x=this.checkp(this.menuplacement[i])
			else this.y=this.checkp(this.menuplacement[i])
		}
		this[i].moveIt(this.x,this.y)
		this[i].bgChange(this.mainbgcoloroff)
		if(!this.menurows) this.y+=this.mainheight+this.checkp(this.pxbetween)
		else this.x+=this.mainwidth+this.checkp(this.pxbetween)
		if(this.submenus[i]!='nosub'){
			this[i].subs=new makeMenu(this,'div'+this.name+'Sub'+i,'',1,i,-1)
			if(!this.menurows) this[i].subs.moveIt(this.subplacement+this[i].x,this[i].y)
			else this[i].subs.moveIt(this[i].x,this[i].y+this.subplacement)
			this.suby=0;
			this[i].sub=new Array()
			for(j=0;j<this.submenus[i]["main"];j++){
				this[i].sub[j]=new makeMenu(this,'div'+this.name+'Sub'+i+'_'+j,'div'+this.name+'Sub'+i,2,i,j)
				this[i].sub[j].clipTo(0,this.subwidth,this.subheight,0,1)
				this[i].sub[j].moveIt(0,this.suby)
				this[i].sub[j].bgChange(this.subbgcoloroff)
				this.suby+=this.subheight
				if(this.submenus[i]["submenus"][j]>0){
					this.subsuby=0
					this[i].sub[j].subs=new makeMenu(this,'div'+this.name+'Sub'+i+'_'+j+'_sub','',1,i,j)
					this[i].sub[j].subs.moveIt(this[i].subs.x+this.subsubXplacement,this[i].subs.y+this[i].sub[j].y+this.subsubYplacement)
					this[i].sub[j].sub=new Array()
					for(a=0;a<this.submenus[i]["submenus"][j];a++){
						this[i].sub[j].sub[a]=new makeMenu(this,'div'+this.name+'Sub'+i+'_'+j+'_sub'+a,'div'+this.name+'Sub'+i+'_'+j+'_sub',3,i,j,a)
						this[i].sub[j].sub[a].clipTo(0,this.subsubwidth,this.subsubheight,0,1)
						this[i].sub[j].sub[a].moveIt(0,this.subsuby)
						this[i].sub[j].sub[a].bgChange(this.subsubbgcoloroff)
						this.subsuby+=this.subsubheight
					}
					this[i].sub[j].subs.clipTo(0,this.subsubwidth,0,0,1)
					this[i].sub[j].subs.clipheight=this.subsuby
				}else this[i].sub[j].subs=0
			}
			this[i].subs.clipTo(0,this.subwidth,0,0,1)
			this[i].subs.clipheight=this.suby
		}else this[i].subs=0
	}
	setTimeout("window.onresize=resized;",500)
	if(this.menueventoff=="mouse"){
		explorerev+=this.name+".hidemain(-1);"
		document.onmouseover=new Function(explorerev)
	}
}
function resized(){
	page2=new makePageCoords()
	if(page2.x2!=page.x2 || page.y2!=page2.y2) location.reload()
}

/*********************************************************************************************
Mouseevents (name==this (as in made object, not the event "this"))
*********************************************************************************************/
function cancelEv(){
	if(bw.ie4 || bw.ie5) window.event.cancelBubble=true
}
function mmover(num,name){
	name[num].bgChange(name.mainbgcoloron)
	if(name.menueventon=="mouse") name.menumain(num,1)
	name[num].nssubover=true
	cancelEv()
}
function mmout(num,name){
	if(!isNaN(num)){
		if(name[num].subs==0 || !name.stayoncolor || !name[num].active)
		name[num].bgChange(name.mainbgcoloroff); 
		name[num].nssubover=false
		if(name.menueventoff=="mouse") if(bw.ns4) setTimeout("if(!"+name.name+"["+num+"].nssubover) "+name.name+".hideactive("+num+")",100)
	} 
	cancelEv()
}
function submmover(num,subnum,name){
	name[num].sub[subnum].bgChange(name.subbgcoloron)
	if(name.menueventon=="mouse") {name.menusub(num,subnum,1)}
	name[num].nssubover=true
	cancelEv()
}
function submmout(num,subnum,name){
	if(!isNaN(subnum)){
		name[num].nssubover=false;
		if(!name.stayoncolor || !name[num].sub[subnum].active || name[num].sub[subnum].subs==0)
		name[num].sub[subnum].bgChange(name.subbgcoloroff)
	}
	cancelEv()
}
function subsubmmover(num,subnum,subsubnum,name){
	if(!isNaN(subnum)){
		name[num].sub[subnum].sub[subsubnum].bgChange(name.subsubbgcoloron); 
		name[num].nssubover=true
	}
	cancelEv()
}
function subsubmmout(num,subnum,subsubnum,name){
	if(!isNaN(subnum)){
		name[num].nssubover=false; 
		name[num].sub[subnum].sub[subsubnum].bgChange(name.subsubbgcoloroff)
	}
	cancelEv()
}
/*********************************************************************************************
Showing submenus
*********************************************************************************************/
function menumain(num,mouse){
	if(this[num].subs!=0){
		clearTimeout(this[num].subs.tim)
		if(this[num].subs.clipy==0 || mouse){
			this.hidemain(num); this[num].subs.clipOut(this.menuspeed); this[num].active=1
		}else{
			this.hidemain(-1); this[num].active=0
		}
	}
	else{
		this.hidemain(-1);
		this[num].bgChange(this.mainbgcoloron,this.mainHilite)
	}
}
/*********************************************************************************************
Showing subsubmenus
*********************************************************************************************/
function menusub(num,sub,mouse){
	this.hidesubs(num,sub)
	if(this[num].sub[sub].subs!=0){
		if(this[num].sub[sub].subs.clipy==0 || mouse){
			this[num].sub[sub].active=1
			this[num].sub[sub].subs.clipOut(this.menusubspeed)
		}else{
			this[num].sub[sub].active=0
			this[num].sub[sub].subs.clipIn(this.menusubspeed)
		}
	}
}
/*********************************************************************************************
Hides the other sub menuitems if any are shown. Also calls the hidesubs to hide any showing
submenus.
*********************************************************************************************/
function hidemain(num){
	for(i=0;i<this.mainmenus;i++){
		if(this[i].subs!=0){
			if(this[i].subs.clipy<=this[i].subs.clipheight){
				this.hidesubs(i,100)
				if(i!=num){
					clearTimeout(this[i].subs.tim)
					this[i].active=0
					this[i].bgChange(this.mainbgcoloroff)
					if(this.menurows)this[i].subs.clipIn(this.menuspeed)
					else{this[i].subs.clipy=0; this[i].subs.clipTo(0,this[i].subs.clipx,this[i].subs.clipy,0,1)}
				}
			}
		}else this[i].bgChange(this.mainbgcoloroff)
	}
}
/*********************************************************************************************
Hides the active submenuitems
*********************************************************************************************/
function hideactive(num){
	if(this[num].subs!=0){
		this.hidesubs(num,100)
		clearTimeout(this[num].subs.tim)
		this[num].active=0
		this[num].bgChange(this.mainbgcoloroff)
		if(this.menurows)this[num].subs.clipIn(this.menuspeed)
		else{this[num].subs.clipy=0; this[num].subs.clipTo(0,this[num].subs.clipx,this[num].subs.clipy,0,1)}
	}
}
/*********************************************************************************************
Hides the other subsub menuitems if any are shown.
*********************************************************************************************/
function hidesubs(num,sub){
	for(j=0;j<this[num].sub.length;j++){
		if(this[num].sub[j].subs!=0 && j!=sub){
			if(this[num].sub[j].subs.clipy<=this[num].sub[j].subs.clipy
			|| this[num].subs.clipy<this[num].subs.clipheight){
				clearTimeout(this[num].sub[j].subs.tim)
				this[num].sub[j].active=0
				this[num].sub[j].bgChange(this.subbgcoloroff)
				this[num].sub[j].subs.clipy=0
				this[num].sub[j].subs.clipTo(0,this[num].sub[j].subs.clipx,this[num].sub[j].subs.clipy,0,1)
			}
		}
	}
}
/*********************************************************************************************
These are the functions that writes the style and menus to the page. 
*********************************************************************************************/
function makeStyle(){
	str='\n<style type="text/css">\n'
	str+="\n<!-- DHTML CoolMenus from www.bratta.com -->\n\n"
	str+='\tDIV.cl'+this.name+'Main{position:absolute; z-index:51; clip:rect(0,0,0,0); overflow:hidden; width:'+(this.mainwidth-10)+'; '+this.clMain+'}\n'
	str+='\tDIV.cl'+this.name+'Sub{position:absolute; z-index:52; clip:rect(0,0,0,0); overflow:hidden; width:'+(this.subwidth-10)+'; '+this.clSub+'}\n'
	str+='\tDIV.cl'+this.name+'SubSub{position:absolute; z-index:54; clip:rect(0,0,0,0); width:'+(this.subsubwidth-10)+'; '+this.clSubSub+'}\n'
	str+='\tDIV.cl'+this.name+'Subs{position:absolute; z-index:53; clip:rect(0,0,0,0); overflow:hidden}\n'
	str+='\t#div'+this.name+'Backgroundbar{position:absolute; z-index:50; clip:rect(0,0,0,0); overflow:hidden}\n'
	str+='\tA.clA'+this.name+'Main{'+this.clAMain+'}\n'
	str+='\tA.clA'+this.name+'Sub{'+this.clASub+'}\n'
	str+='\tA.clA'+this.name+'SubSub{'+this.clASubSub+'}\n'
	str+='</style>\n\n'
	document.write(str)
}
function makeMain(num,text,link,target){
	str=""
	if(this.backgroundbar && num==0){str+='\n<div id="div'+this.name+'Backgroundbar"></div>\n'}
	str+='<div id="div'+this.name+'Main'+num+'" class="cl'+this.name+'Main">'
	if(link){ 
		link=(link=='#' || (link.indexOf(':')>-1))?link:lvl+link;
		str+='<a href="'+link+'"'; this.submenus[num]='nosub'
	}
	else str+='<a href="#" onclick="'+this.name+'.menumain('+num+'); return false"'
	if(target) str+=' target="'+target+'" '
	str+=' class="clA'+this.name+'Main">'+text+'</a></div>\n'
	this.mainmenus++; 
	document.write(str)
}
function makeSub(num,subnum,text,link,total,target){
	str=""
	if(subnum==0) str='<div id="div'+this.name+'Sub'+num+'" class="cl'+this.name+'Subs">\n'
	str+='\t<div id="div'+this.name+'Sub'+num+'_'+subnum+'" class="cl'+this.name+'Sub">'
	if(link){
		link=(link=='#' || (link.indexOf(':')>-1))?link:lvl+link;
		str+='<a href="'+link+'"';
	}
	else str+='<a href="#" onclick="'+this.name+'.menusub('+num+','+subnum+'); return false"'
	if(target) str+=' target="'+target+'" '
	str+=' class="clA'+this.name+'Sub">'+text+'</a></div>\n'
	if(subnum==total-1){
		str+='</div>\n'; this.submenus[num]=new Array()
		this.submenus[num]["main"]=total; this.submenus[num]["submenus"]=new Array()
	}
	document.write(str)
}
function makeSubSub(num,subnum,subsubnum,text,link,total,target){
	str=""
	if(subsubnum==0) str='<div id="div'+this.name+'Sub'+num+'_'+subnum+'_sub" class="cl'+this.name+'Subs">\n'
	str+='\t<div id="div'+this.name+'Sub'+num+'_'+subnum+'_sub'+subsubnum+'" class="cl'+this.name+'SubSub">'
	if(link){ 
		link=(link=='#' || (link.indexOf(':')>-1))?link:lvl+link;
		str+='<a href="'+link+'"';
	} 
	else str+='<a href="#"'
	if(target) str+=' target="'+target+'" '
	str+=' class="clA'+this.name+'SubSub">'+text+'</a></div>\n'
	if(subsubnum==total-1){str+='</div>\n'; this.submenus[num]["submenus"][subnum]=total}
	document.write(str)
}
/*********************************************************************************************
END Menu script
*********************************************************************************************/
