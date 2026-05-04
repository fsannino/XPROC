//*** Default browsercheck, added to all scripts!
function checkBrowser()
{
	this.ver=navigator.appVersion
	this.dom=document.getElementById?1:0
	//*** this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom)?1:0; é o que tinha antes
	//*** alterei a linha abaixo para checagem tb do IE 6 Beta :-) --> 
	this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom || this.ver.indexOf("MSIE 6.0b")>-1 || this.ver.indexOf("MSIE 6")>-1 && this.dom)?1:0;
	this.ie4=(document.all && !this.dom)?1:0;
	this.ns5=(this.dom && parseInt(this.ver) >= 5) ?1:0;
	this.ns4=(document.layers && !this.dom)?1:0;
	this.bw=(this.ie5 || this.ie4 || this.ns4 || this.ns5)
	return this
}
var bw=new checkBrowser()
var qtd=0
var posicao
var posicao2
var raiz=''
//raiz = 'xproc'

var a = document.URL;
	var n=0;

	for (var i = 1 ; i < 1000; i++)
	{
	var final=a.slice(0,i)
	var t=a.slice(i-1,i);
	if (t=='/')
	{
	n = n + 1;
	}
	if(n == 4)
	{
	i = 1000;
	}
	}
var tam=final.length;
raiz = final.slice(0,tam-1);
//raiz = 'http://s6000ws10.corp.petrobras.biz/xproc'	
		
//Ie var
var explorerev=''

var lvl=''
var offline= '';
var online= 'http://www.aforbes.com.br';
var cnt=0;
var off_cnt=0

function getLevel()
{
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

function goMenus()
{
	/********************************************************************************
	Variables to set.
	
	Se lembre isso para fixar para fontsize e para fonttype o jogo isso no stylesheet  
	sobre!
	********************************************************************************/	

	//Fazendo um objeto do menu
	oMenu=new menuObj('oMenu') //Coloque um nome para o menu. Deve ser unico para cada menu
	//Configuração das variáveis do objeto menu
	
	var intAltura;
	var intLargura;
	
	var intSubheight;
	var intSubwidth;
	
	var intSubSubheight;
	var intSubSubwidth;
	
	var intSubSubYplacement;
	var intSubSubXplacement;
		
	//*** DEFINE O TRAMANHO DO MENU
	switch (str_CategoriaUsuario)
	{
		case 'indexB.js':
			intAltura = 22;
			intLargura = 120;
			intSubheight = 25;
			intSubwidth = 150;
			intSubSubheight = 25;
			intSubSubwidth = 150;
			intSubSubYplacement = 5; 
			intSubSubXplacement = 120;
			break;
			
		case 'indexC.js':
			intAltura = 22;
			intLargura = 120;
			intSubheight = 25;
			intSubwidth = 150;
			intSubSubheight = 25;
			intSubSubwidth = 150;
			intSubSubYplacement = 5; 
			intSubSubXplacement = 120;
			break;

		case 'indexD.js':
			intAltura = 22;
			intLargura = 120;
			intSubheight = 25;
			intSubwidth = 150;
			intSubSubheight = 25;
			intSubSubwidth = 150;
			intSubSubYplacement = 5; 
			intSubSubXplacement = 120;
			break;
		
		case 'indexE.js': 
			intAltura = 22;
			intLargura = 120;
			intSubheight = 25;
			intSubwidth = 150;
			intSubSubheight = 25;
			intSubSubwidth = 160;
			intSubSubYplacement = 5; 
			intSubSubXplacement = 120;
			break;
		
		case 'indexQ.js': 
			intAltura = 22;
			intLargura = 70;
			intSubheight = 25;
			intSubwidth = 140;
			intSubSubheight = 25;
			intSubSubwidth = 160;
			intSubSubYplacement = 5; 
			intSubSubXplacement = 80;
			break;
		
		case 'indexZ.js': 
			intAltura = 22;
			intLargura = 120;
			intSubheight = 25;
			intSubwidth = 150;
			intSubSubheight = 25;
			intSubSubwidth = 160;
			intSubSubYplacement = 5; 
			intSubSubXplacement = 120;
			break;
		
		default:
			intAltura = 22;
			intLargura = 80;
			intSubheight = 25;
			intSubwidth = 120;
			intSubSubheight = 25;
			intSubSubwidth = 150;
			intSubSubYplacement = 5; 
			intSubSubXplacement = 120;
			break;
	}	
	
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
	oMenu.mainheight = intAltura //A altura do menuitems principal em pixel ou%
	oMenu.mainwidth = intLargura //A largura do menuitems principal em pixel ou%
	
	/*Estas são variáveis novas. Neste exemplo eles são fixos como a versão prévia*/
	//oMenu.subwidth=oMenu.mainwidth // ** NEW ** A largura dos submenu (largura igual a do menu)
	//oMenu.subheight=oMenu.mainheight // ** NEW ** A altura dos submenu (largura igual a do menu)
	oMenu.subheight=intSubheight //Caso você não queira a altura igual a do menu
	oMenu.subwidth=intSubwidth //Caso você não queira a largura igual a do menu
		
	//oMenu.subsubwidth=oMenu.mainwidth // ** NEW ** A largura do subsubmenus em pixel ou% 
	//oMenu.subsubheight=oMenu.subheight //** NEW ** A altura se o subsubitems em pixel ou% 
	oMenu.subsubheight = intSubSubheight //Caso você não queira a altura igual a do submenu
	oMenu.subsubwidth = intSubSubwidth //Caso você não queira a largura igual a do submenu
	
	//Escrevendo fora o estilo para o menu (deixe esta linha!)
	oMenu.makeStyle()
	
	oMenu.subplacement=oMenu.mainheight //** NEW ** altura dos menus que irão aparecer
	//oMenu.subsubXplacement=oMenu.subwidth/2 //** NEW ** distância dos submenus em relação aos links baseada nos submenu
	oMenu.subsubYplacement=intSubSubYplacement //** NEW ** Altura dos subsubmenus em relação ao links do submenu
	oMenu.subsubXplacement=intSubSubXplacement //** NEW ** distância dos submenus em relação aos links
	
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
	
	switch (str_CategoriaUsuario)
	{
		case 'indexA.js':	
			oMenu.menuplacement=new Array('0','80','160','240','320','400'); //os primeiros 3 números são a posição horizontal dos menus
			break;
			
		case 'indexB.js':
			oMenu.menuplacement=new Array('0','120','240','360','480','580'); //*** Posição horizontal dos menus
			break;
			
		case 'indexC.js':
			oMenu.menuplacement=new Array('0','120'); //*** Posição horizontal dos menus
			break;
			
		case 'indexD.js':	
			oMenu.menuplacement=new Array('0'); //*** Posição horizontal dos menus
			break;
			
		case 'indexE.js':	
			oMenu.menuplacement=new Array('0','120'); //*** Posição horizontal dos menus
			break;
			
		case 'indexF.js':	
			oMenu.menuplacement=new Array('0','80','160','240','320','400','480'); //*** Posição horizontal dos menus
			break;
		
		case 'indexQ.js':	
			oMenu.menuplacement=new Array('0','70','140','210','280','350','420','490','560','630','700'); //*** Posição horizontal dos menus
			break;		
		
		case 'indexV.js':	
			oMenu.menuplacement=new Array('0','80','160','240','320','400','480','560'); //*** Posição horizontal dos menus
			break;			
		
		case 'indexZ.js':	
			oMenu.menuplacement=new Array('0','120'); //*** Posição horizontal dos menus
			break;
			
		default:
			oMenu.menuplacement=new Array('0','80','160'); //*** Posição horizontal dos menus
			break;
	}	
	
	//Se você usa o "direito ao lado de eachother" você hipocrisia que quanto pixel deveriam estar entre cada aqui
	oMenu.pxbetween=0 //in pixel or %
	
	//E você pode fixar onde deveria começar da esquerda aqui
	oMenu.fromleft=45 //in pixel or %
	
	//Altura do menu em relação ao topo do browser
	oMenu.fromtop=80 //in pixel or %
	
	/********************************************************************************
	********************	Construindo os menus  ********************************
	********************************************************************************/
	
	switch (str_CategoriaUsuario)
	{
		case 'indexA.js':		
			processo(0); //*** REFERENTE AO PROCESSO		
			cenario(1);	//*** REFERENTE AO CENÁRIO				
			funcao(2); //*** REFERENTE AO FUNÇÃO		
			desenho(3);	//*** REFERENTE AO DESENHO
			//cursos(4);	//*** REFERENTE AO CURSOS		
			consulta(4); //*** REFERENTE AO CONSULTA
			break;
		
		case 'indexB.js':		
			cadastro(0); //*** REFERENTE AO CADASTRO				
			decomposicao(1); //*** REFERENTE AO DECOMPOSIÇÃO		
			consulta(2); //*** REFERENTE AO CONSULTA		
			usuario(3); //*** REFERENTE AO USUÁRIO		
			escopo(4); //*** REFERENTE AO ESCOPO		
			cenario_Simples(5); //*** REFERENTE AO CENÁRIO
			break;
			
		case 'indexC.js':	
			consulta(0); //*** REFERENTE AO CONSULTA		
			escopo_IndexC(1); //*** REFERENTE AO ESCOPO
			break;	
			
		case 'indexD.js':		
			consulta(0); //*** REFERENTE AO CONSULTA	
			break;
	
		case 'indexE.js':
			//cursos(0); //*** REFERENTE AO CURSOS		
			consulta(0); //*** REFERENTE AO CONSULTA
			break;
			
		case 'indexF.js':				
			processo(0); //*** REFERENTE AO PROCESSO		
			cenario(1); //*** REFERENTE AO CENÁRIO		
			funcao_IndexF(2); //*** REFERENTE A FUNÇÃO		
			perfil(3);	//*** REFERENTE A PERFIL				
			desenho(4); //*** REFERENTE A DESENHO				
			//cursos(5); //*** REFERENTE A CURSOS				
			consulta(5); //*** REFERENTE A CONSULTAS
			break;
		
		case 'indexG.js':
			perfil_IndexG(0); //*** REFERENTE A PERFIL		
			consulta(1); //*** REFERENTE A CONSULTAS		
			break;
	
		case 'indexH.js':
			perfil_IndexH(0); //*** REFERENTE A PERFIL		
			goLive(1); //*** REFERENTE A GOLIVE
			consulta(2); //*** REFERENTE A CONSULTAS		
			break;
		
		case 'indexP.js':
			funcao_IndexP(0); //*** REFERENTE A FUNÇÃO		
			perfil_IndexP(1); //*** REFERENTE A PERFIL	
			consulta(2);	//*** REFERENTE A CONSULTAS				
			break;
	
		case 'indexQ.js':
			escopo_IndexQ(0); //*** REFERENTE AO ESCOPO		
			processo(1); //*** REFERENTE AO PROCESSO		
			cenario_Composto(2); //*** REFERENTE AO CENÁRIO		
			funcao_IndexQ(3); //*** REFERENTE AO FUNÇÃO		
			perfil_IndexQ(4); //*** REFERENTE AO PERFIL		
			desenho_IndexQ(5); //*** REFERENTE AO DESENHO		
			cursos_IndexQ(6); //*** REFERENTE AO CURSOS		
			cases_IndexQ(7); //*** REFERENTE AO CASES		
			usuario(8); //*** REFERENTE AO USUÁRIO		
			pep_IndexQ(9); //*** REFERENTE AO PEP		
			consulta(10); //*** REFERENTE A CONSULTA		
			break;
	
		case 'indexV.js':
			processo(0); //*** REFERENTE AO PROCESSO		
			cenario(1); //*** REFERENTE AO CENÁRIO		
			funcaoV(2); //*** REFERENTE A FUNÇÃO		
			perfil_IndexV(3); //*** REFERENTE AO PERFIL	
			desenho(4); //*** REFERENTE AO DESENHO		
			//cursos(5); //*** REFERENTE AOS CURSOS		
			goLive_IndexV(5); //*** REFERENTE A GOLIVE
			consulta(6); //*** REFERENTE A CONSULTA		
			break;
			
		case 'indexZ.js':
			cursos_IndexZ(0); //*** REFERENTE AO CURSOS		
			consulta(1); //*** REFERENTE AO CONSULTA
			break;

		case 'indexW.js':				
			cursos(0); //*** REFERENTE A CURSOS				
			consulta(1); //*** REFERENTE A CONSULTAS
			break;
			
		default:
			consulta(0); //*** REFERENTE AO CONSULTA	
			break;
	}
	
	/********************************************************************************
	End menu construction
	********************************************************************************/
			
	//When all the menus are written out we initiates the menu
	oMenu.construct()
}

/********************************************************************************
Object constructor and object functions
********************************************************************************/
function makePageCoords()
{
	this.x=0;this.x2=(bw.ns4 || bw.ns5)?innerWidth:document.body.offsetWidth-20;
	this.y=0;this.y2=(bw.ns4 || bw.ns5)?innerHeight:document.body.offsetHeight-5;
	this.x50=this.x2/2;	this.y50=this.y2/2;
	return this;
}

function makeMenu(parent,obj,nest,type,num,subnum,subsubnum)
{
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

function b_clipIn(speed)
{
	if(this.clipy>0){
		this.clipy-=speed
		if(this.clipy<0) this.clipy=0
		this.clipTo(0,this.clipx,this.clipy,0,1)
		this.tim=setTimeout(this.obj+".clipIn("+speed+")",10)
	}else{this.clipy=0; this.clipTo(0,this.clipx,this.clipy,0,1)}	
}

function b_clipOut(speed)
{
	if(this.clipy<this.clipheight){
		this.clipy+=speed
		this.clipTo(0,this.clipx,this.clipy,0,1)
		this.tim=setTimeout(this.obj+".clipOut("+speed+")",10)
	}else{this.clipy=this.clipheight; this.clipTo(0,this.clipx,this.clipy,0,1)}
}

//*** Page variable, holds the width and height of the document. (see documentsize tutorial on bratta.com/dhtml)
var page=new makePageCoords()

/********************************************************************************
Checking if the values are % or not.
********************************************************************************/
function checkp(num,lefttop)
{
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
function menuObj(name)
{
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

function constructMenu()
{
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

function resized()
{
	page2=new makePageCoords()
	if(page2.x2!=page.x2 || page.y2!=page2.y2) location.reload()
}

/*********************************************************************************************
Mouseevents (name==this (as in made object, not the event "this"))
*********************************************************************************************/
function cancelEv()
{
	if(bw.ie4 || bw.ie5) window.event.cancelBubble=true
}

function mmover(num,name)
{
	name[num].bgChange(name.mainbgcoloron)
	if(name.menueventon=="mouse") name.menumain(num,1)
	name[num].nssubover=true
	cancelEv()
}

function mmout(num,name)
{
	if(!isNaN(num)){
		if(name[num].subs==0 || !name.stayoncolor || !name[num].active)
		name[num].bgChange(name.mainbgcoloroff); 
		name[num].nssubover=false
		if(name.menueventoff=="mouse") if(bw.ns4) setTimeout("if(!"+name.name+"["+num+"].nssubover) "+name.name+".hideactive("+num+")",100)
	} 
	cancelEv()
}

function submmover(num,subnum,name)
{
	name[num].sub[subnum].bgChange(name.subbgcoloron)
	if(name.menueventon=="mouse") {name.menusub(num,subnum,1)}
	name[num].nssubover=true
	cancelEv()
}

function submmout(num,subnum,name)
{
	if(!isNaN(subnum)){
		name[num].nssubover=false;
		if(!name.stayoncolor || !name[num].sub[subnum].active || name[num].sub[subnum].subs==0)
		name[num].sub[subnum].bgChange(name.subbgcoloroff)
	}
	cancelEv()
}

function subsubmmover(num,subnum,subsubnum,name)
{
	if(!isNaN(subnum)){
		name[num].sub[subnum].sub[subsubnum].bgChange(name.subsubbgcoloron); 
		name[num].nssubover=true
	}
	cancelEv()
}

function subsubmmout(num,subnum,subsubnum,name)
{
	if(!isNaN(subnum)){
		name[num].nssubover=false; 
		name[num].sub[subnum].sub[subsubnum].bgChange(name.subsubbgcoloroff)
	}
	cancelEv()
}

/*********************************************************************************************
Showing submenus
*********************************************************************************************/
function menumain(num,mouse)
{
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
function menusub(num,sub,mouse)
{
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
function hidemain(num)
{
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
function hideactive(num)
{
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
function hidesubs(num,sub)
{
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
function makeStyle()
{
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

function makeMain(num,text,link,target)
{
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

function makeSub(num,subnum,text,link,total,target)
{
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

function makeSubSub(num,subnum,subsubnum,text,link,total,target)
{
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


//////////////////////  FUNÇÕES DO MENU  ///////////////////////////////////////////////////////////////////	

function processo(posicao)
{	//*** IndexA.js	
	oMenu.makeMain(posicao,'PROCESSO',0)
	
	qtd_1=3;
	oMenu.makeSub(posicao,0,'Processo',0,qtd_1)
	oMenu.makeSub(posicao,1,'Sub-Processo',0,qtd_1)
	oMenu.makeSub(posicao,2,'Decomposição',0,qtd_1)
	
	qtd_10=3;
	oMenu.makeSubSub(posicao,0,0,'Novo',raiz + '/asp/cadas_processo.asp?txtOpc=1',qtd_10)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/altera_processo.asp',qtd_10)
	oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/exclui_.asp',qtd_10)
	
    qtd_11=3;
	oMenu.makeSubSub(posicao,1,0,'Novo',raiz + '/asp/cadas_sub_processo.asp?txtOpc=1',qtd_11)
	oMenu.makeSubSub(posicao,1,1,'Alterar',raiz + '/asp/altera_sub_processo.asp?txtOpc=1',qtd_11)
	oMenu.makeSubSub(posicao,1,2,'Excluir',raiz + '/asp/exclui.asp',qtd_11)
	
	qtd_12=1;
	oMenu.makeSubSub(posicao,2,0,'Por Sub-Processo',raiz + '/asp/selec_Mega_Proc_Sub_Processo.asp?txtOpc=1',qtd_12)
}

function cenario(posicao)
{	//*** IndexA.js	
	oMenu.makeMain(posicao,'CENÁRIO',0)
	
	qtd_2=2;
	oMenu.makeSub(posicao,0,'Cenário',0,qtd_2)	
	oMenu.makeSub(posicao,1,'Classe',0,qtd_2)
	//oMenu.makeSub(posicao,2,'Solicita Escopo','http://164.85.62.152/ComSinergia/DataBase33.nsf/frmChamado?OpenForm',qtd_2)
	//oMenu.makeSub(posicao,3,'Consulta Solicitação','http://164.85.62.152/ComSinergia/DataBase33.nsf/1D503FF15BA7EBF683256D12004C6ECD?OpenPage',qtd_2)

	qtd_20=8;
	oMenu.makeSubSub(posicao,0,0,'Novo',raiz + '/asp/cenario/cad_cenario.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/cenario/altera_cenario.asp',qtd_20)
	//oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/cenario/excluir_cenario.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,2,'Copia Cenario Onda',raiz + '/asp/cenario/copia_cenario_onda.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,3,'Editar Cenário/Transação',raiz + '/asp/cenario/selec_Mega_Proc_Sub_Cenario2.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,4,'Editar Ordem Cenário',raiz + '/asp/cenario/sel_cenario_altera_sequencia.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,5,'Alterar Assunto em massa', raiz + '/asp/cenario/sel_cenario_altera_assunto.asp?txtOPT=1',qtd_20)
	oMenu.makeSubSub(posicao,0,6,'Cadastro Pacote de Teste', raiz + '/asp/cenario/cad_hist_pacote.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,7,'Definir Fluxo Mestre',raiz + '/asp/cenario/sel_cenario_gera_fluxo.asp',qtd_20)
	
	qtd_21=1;	
	oMenu.makeSubSub(posicao,1,0,'Alterar Classe',raiz + '/asp/cenario/alterar_classe.asp',qtd_21)
}

function funcao(posicao)
{	//*** IndexA.js	
	oMenu.makeMain(posicao,'FUNÇÃO',0)
	
	qtd_3=8;
	//oMenu.makeSub(posicao,0,'Função',0,qtd_3)
	oMenu.makeSub(posicao,0,'Função Transações',0,qtd_3)
	oMenu.makeSub(posicao,1,'Ag.Ativ x Ativ x Tran',raiz + '/asp/relacao_master_.asp',qtd_3)
	oMenu.makeSub(posicao,2,'Função Conflitante',raiz + '/asp/funcao/func_confl.asp?',qtd_3)
	oMenu.makeSub(posicao,3,'Função x Assunto',0,qtd_3)
	//oMenu.makeSub(posicao,4,'Orient Gerais',0,qtd_3)
	//oMenu.makeSub(posicao,5,'Termos Gerais',0,qtd_3)
	oMenu.makeSub(posicao,4,'Orient Mega',0,qtd_3)
	oMenu.makeSub(posicao,5,'Termos Mega ',0,qtd_3)
	oMenu.makeSub(posicao,6,'Assuntos Mega ',0,qtd_3)
	oMenu.makeSub(posicao,7,'Orie.Funcao(OBS)',raiz + '/asp/orie_mape/seleciona_funcao.asp?pOpt=CF',qtd_3)

	qtd_30=3;
	//oMenu.makeSubSub(posicao,0,0,'Nova',raiz + '/asp/funcao/cad_funcao.asp',qtd_30)
	//oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=1',qtd_30)
	//oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=2',qtd_30)
	
	qtd_30=2;
	posicao2 = 0
	oMenu.makeSubSub(posicao,posicao2,0,'Fun-Tra (mesmo MEGA)',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=3',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,1,'Fun-Tra (outro MEGA)',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=4',qtd_30)
	
	qtd_30=1;	
	posicao2 = 3
	oMenu.makeSubSub(posicao,posicao2,0,'Altera assunto em massa',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=6',qtd_30)

	//qtd_30=3;	
	//posicao2 = 0
	//oMenu.makeSubSub(posicao,posicao2,0,'Inclui Orientações',raiz + '/asp/orie_mape/inclui_ori_gerais_mapeamento.asp',qtd_30)
	//oMenu.makeSubSub(posicao,posicao2,1,'Altera Orientações',raiz + '/asp/orie_mape/seleciona_ori_gerais_mapeamento.asp?pOpt=A',qtd_30)
	//oMenu.makeSubSub(posicao,posicao2,2,'Excluir Orientações',raiz + '/asp/orie_mape/seleciona_ori_gerais_mapeamento.asp?pOpt=E',qtd_30)
	
	//qtd_30=3;	
	//posicao2 = 0
	//oMenu.makeSubSub(posicao,posicao2,0,'Inclui Termos',raiz + '/asp/orie_mape/inclui_ori_gerais_mape_termos.asp',qtd_30)
	//oMenu.makeSubSub(posicao,posicao2,1,'Altera Termos',raiz + '/asp/orie_mape/seleciona_ori_gerais_mape_termos.asp?pOpt=A',qtd_30)
	//oMenu.makeSubSub(posicao,posicao2,2,'Excluir Termos',raiz + '/asp/orie_mape/seleciona_ori_gerais_mape_termos.asp?pOpt=E',qtd_30)

	qtd_30=3;
	posicao2 = 4
	oMenu.makeSubSub(posicao,posicao2,0,'Inclui Orientações Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IO',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,1,'Altera Orientações Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_mapeamento.asp?pOpt=AO',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,2,'Excluir Orientações Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_mapeamento.asp?pOpt=EO',qtd_30)
	
	qtd_30=3;
	posicao2 = 5
	oMenu.makeSubSub(posicao,posicao2,0,'Inclui Termos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IT',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,1,'Altera Termos Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_termo.asp?pOpt=AT',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,2,'Excluir Termos Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_termo.asp?pOpt=ET',qtd_30)

	qtd_30=3;
	posicao2 = 6
	oMenu.makeSubSub(posicao,posicao2,0,'Inclui Assuntos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,1,'Altera Assuntos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=AM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,2,'Excluir Assuntos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=EM&pOpt2=M',qtd_30)
}

function funcaoV(posicao)
{	//*** IndexV.js
	oMenu.makeMain(posicao,'FUNÇÃO',0)
	
	qtd_3=11;
	oMenu.makeSub(posicao,0,'Função',0,qtd_3)
	oMenu.makeSub(posicao,1,'Função x Trasações',0,qtd_3)
	oMenu.makeSub(posicao,2,'Ag.Ativ x Ativ x Tran',raiz + '/asp/relacao_master_.asp',qtd_3)
	oMenu.makeSub(posicao,3,'Função Conflitante',raiz + '/asp/funcao/func_confl.asp?',qtd_3)
	oMenu.makeSub(posicao,4,'Função x Assunto',0,qtd_3)
	oMenu.makeSub(posicao,5,'Orient Gerais',0,qtd_3)
	oMenu.makeSub(posicao,6,'Termos Gerais',0,qtd_3)
	oMenu.makeSub(posicao,7,'Orient Mega',0,qtd_3)
	oMenu.makeSub(posicao,8,'Termos Mega ',0,qtd_3)
	oMenu.makeSub(posicao,9,'Assuntos Mega ',0,qtd_3)
	oMenu.makeSub(posicao,10,'Orie.Funcao(OBS)',raiz + '/asp/orie_mape/seleciona_funcao.asp?pOpt=CF',qtd_3)

	qtd_30=3;
	oMenu.makeSubSub(posicao,0,0,'Nova',raiz + '/asp/funcao/cad_funcao.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=1',qtd_30)
	oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=2',qtd_30)
	
	qtd_30=2;
	oMenu.makeSubSub(posicao,1,0,'Fun-Tra (mesmo MEGA)',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=3',qtd_30)
	oMenu.makeSubSub(posicao,1,1,'Fun-Tra (outro MEGA)',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=4',qtd_30)
	
	qtd_30=1;	
	oMenu.makeSubSub(posicao,4,0,'Altera assunto em massa',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=6',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,5,0,'Inclui Orientações',raiz + '/asp/orie_mape/inclui_ori_gerais_mapeamento.asp',qtd_30)
	oMenu.makeSubSub(posicao,5,1,'Altera Orientações',raiz + '/asp/orie_mape/seleciona_ori_gerais_mapeamento.asp?pOpt=A',qtd_30)
	oMenu.makeSubSub(posicao,5,2,'Excluir Orientações',raiz + '/asp/orie_mape/seleciona_ori_gerais_mapeamento.asp?pOpt=E',qtd_30)
	
	qtd_30=3;	
	oMenu.makeSubSub(posicao,6,0,'Inclui Termos',raiz + '/asp/orie_mape/inclui_ori_gerais_mape_termos.asp',qtd_30)
	oMenu.makeSubSub(posicao,6,1,'Altera Termos',raiz + '/asp/orie_mape/seleciona_ori_gerais_mape_termos.asp?pOpt=A',qtd_30)
	oMenu.makeSubSub(posicao,6,2,'Excluir Termos',raiz + '/asp/orie_mape/seleciona_ori_gerais_mape_termos.asp?pOpt=E',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,7,0,'Inclui Orientações Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IO',qtd_30)
	oMenu.makeSubSub(posicao,7,1,'Altera Orientações Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_mapeamento.asp?pOpt=AO',qtd_30)
	oMenu.makeSubSub(posicao,7,2,'Excluir Orientações Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_mapeamento.asp?pOpt=EO',qtd_30)
	
	qtd_30=3;	
	oMenu.makeSubSub(posicao,8,0,'Inclui Termos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IT',qtd_30)
	oMenu.makeSubSub(posicao,8,1,'Altera Termos Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_termo.asp?pOpt=AT',qtd_30)
	oMenu.makeSubSub(posicao,8,2,'Excluir Termos Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_termo.asp?pOpt=ET',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,9,0,'Inclui Assuntos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,9,1,'Altera Assuntos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=AM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,9,2,'Excluir Assuntos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=EM&pOpt2=M',qtd_30)
}

function desenho(posicao)
{	//*** IndexA.js, IndexF e IndexV
	oMenu.makeMain(posicao,'DESENHO',0)
	
	qtd_3=1;
	oMenu.makeSub(posicao,0,'Procedimentos',raiz + '/doc/Procedimento_Fluxo_de_Processo.doc',qtd_3,'_blank')

	//qtd_30=1;
	//oMenu.makeSubSub(posicao,0,0,'Procedimentos',raiz + '/doc/Procedimento_Fluxo_de_Processo.doc',qtd_30)
}

function cursos(posicao)
{	//*** IndexA.js	, IndexE.js, IndexF.js e IndexV.js

	oMenu.makeMain(posicao,'CURSOS',0)
	
	qtd_3=1;
	oMenu.makeSub(posicao,0,'Cursos',0,qtd_3)

	qtd_30=1;
	//oMenu.makeSubSub(posicao,0,0,'Novo',raiz + '/asp/curso/cad_curso.asp',qtd_30)
	//oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/curso/seleciona_curso.asp?option=6',qtd_30)
	//oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/curso/seleciona_curso.asp?option=5',qtd_30)
	//oMenu.makeSubSub(posicao,0,0,'Curso x Função x Trans',raiz + '/asp/curso/seleciona_curso.asp?option=2',qtd_30)
	//oMenu.makeSubSub(posicao,0,3,'Curso x Transação',raiz + '/asp/curso/seleciona_curso.asp?option=1',qtd_30)
	//oMenu.makeSubSub(posicao,0,4,'Curso x Cenário',raiz + '/asp/curso/seleciona_curso.asp?option=3',qtd_30)
	//oMenu.makeSubSub(posicao,0,1,'Pré Requisito',raiz + '/asp/curso/seleciona_curso.asp?option=4',qtd_30)
	oMenu.makeSubSub(posicao,0,0,'Lib.Manual (LM)',raiz + '/asp/treinamento/seleciona_lm.asp',qtd_30)	
}

function cursos_IndexZ(posicao)
{	
	oMenu.makeMain(posicao,'CURSOS',0)
	
	qtd_3=1;
	oMenu.makeSub(posicao,0,'Cursos',0,qtd_3)

	qtd_30=5;
	oMenu.makeSubSub(posicao,0,0,'Novo',raiz + '/asp/curso/cad_curso.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/curso/seleciona_curso.asp?option=6',qtd_30)
	//oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/curso/seleciona_curso.asp?option=5',qtd_30)
	oMenu.makeSubSub(posicao,0,2,'Curso x Função x Trans',raiz + '/asp/curso/seleciona_curso.asp?option=2',qtd_30)
	//oMenu.makeSubSub(posicao,0,3,'Curso x Transação',raiz + '/asp/curso/seleciona_curso.asp?option=1',qtd_30)
	//oMenu.makeSubSub(posicao,0,4,'Curso x Cenário',raiz + '/asp/curso/seleciona_curso.asp?option=3',qtd_30)
	oMenu.makeSubSub(posicao,0,3,'Pré Requisito',raiz + '/asp/curso/seleciona_curso.asp?option=4',qtd_30)
	oMenu.makeSubSub(posicao,0,4,'Curso x Correlato',raiz + '/asp/curso/seleciona_curso.asp?option=8',qtd_30)
}

function cadastro(posicao)
{	//*** IndexB.js	
	oMenu.makeMain(posicao,'CADASTRO',0)
	
	qtd_62 = 2
	oMenu.makeSub(posicao,0,'Processo......................',0,qtd_62)
	oMenu.makeSub(posicao,1,'Sub-Processo..................',0,qtd_62)
	
	pos_Sub = 0	
	qtd_62 = 3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cadas_processo.asp?txtOpc=1',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar...................',raiz + '/asp/altera_processo.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Excluir...................',raiz + '/asp/exclui_.asp',qtd_62)
	
	pos_Sub = 1
	qtd_62 = 3	
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cadas_sub_processo.asp?txtOpc=1',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar...................',raiz + '/asp/altera_sub_processo.asp?txtOpc=1',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Excluir...................',raiz + '/asp/exclui.asp',qtd_62)
}

function decomposicao(posicao)
{	//*** IndexB.js		
	oMenu.makeMain(posicao,'DECOMPOSIÇÃO',0)
	
	qtd_62 = 1
	//oMenu.makeSub(posicao,0,'Por Atividade',raiz + '/asp/form_relaciona_ativ_trans4.asp?txtOpc=1',qtd_62)	
	oMenu.makeSub(posicao,0,'Por Sub-Processo',raiz + '/asp/selec_Mega_Proc_Sub_Processo.asp?txtOpc=1',qtd_62)
}

function usuario(posicao)
{	//*** IndexB.js	e IndexQ.js
	oMenu.makeMain(posicao,'USUÁRIO',0)
	
	qtd_7=2;
	oMenu.makeSub(posicao,0,'Cadastro',raiz + '/asp/cad_usuario.asp',qtd_7)
	oMenu.makeSub(posicao,1,'Acesso',raiz + '/asp/cadas_acesso.asp',qtd_7)
}

function escopo(posicao)
{	//*** IndexB.js	

	oMenu.makeMain(posicao,'ESCOPO',0)
	oMenu.makeSub(posicao,0,'Agrup.(Mstr List R3)',0,7)
	oMenu.makeSub(posicao,1,'Atividade',0,7)
	oMenu.makeSub(posicao,2,'Transação',0,7)
	oMenu.makeSub(posicao,3,'Empresa/Unidade',0,7)
	oMenu.makeSub(posicao,4,'Classe',0,7)
	oMenu.makeSub(posicao,5,'Agr.Ativ x Ativ x Transação',raiz + '/asp/relacao_master_.asp',7)
	oMenu.makeSub(posicao,6,'Atividade x Empresa',raiz + '/asp/relacao_ativ_emp_.asp',7)
	
	pos_Sub = 0
	qtd_62=3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cad_modulo.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar...................',raiz + '/asp/altera_modulo.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Excluir...................',raiz + '/asp/exclusao.asp?ID=1',qtd_62)

	pos_Sub = 1
	qtd_62=3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cad_atividade.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar...................',raiz + '/asp/altera_atividade.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Excluir...................',raiz + '/asp/exclusao.asp?ID=2',qtd_62)

	pos_Sub = 2
	qtd_62=3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cad_transacao.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar...................',raiz + '/asp/altera_transacao.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Excluir...................',raiz + '/asp/exclusao.asp?ID=3',qtd_62)

	pos_Sub = 3
	qtd_62=3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cad_empresa.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar...................',raiz + '/asp/altera_empresa.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Excluir...................',raiz + '/asp/exclusao.asp?ID=4',qtd_62)

	pos_Sub = 4
	qtd_62=1
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cenario/cad_classe_mega.asp',qtd_62)
}


function cenario_Simples(posicao)
{
	//*** IndexB.js
		
	oMenu.makeMain(posicao,'CENÁRIO',0)

    oMenu.makeSub(posicao,0,'Cenário',0,1)

	pos_Sub = 0
	qtd_62=3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cenario/cad_cenario.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar Dados..........',raiz + '/asp/cenario/altera_Cenario.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Editar...................',raiz + '/asp/cenario/selec_Mega_Proc_Sub_Cenario.asp',qtd_62)
}

function escopo_IndexC(posicao)
{	//*** IndexC.js

	oMenu.makeMain(posicao,'ESCOPO',0)
	
	qtd_62=6
	oMenu.makeSub(posicao,0,'Agrup.(Mstr List R3)',0,qtd_62)
	oMenu.makeSub(posicao,1,'Atividade',0,qtd_62)
	oMenu.makeSub(posicao,2,'Transação',0,qtd_62)
	oMenu.makeSub(posicao,3,'Empresa/Unidade',0,qtd_62)
	oMenu.makeSub(posicao,4,'Agr.Ativ x Ativ x Transação',raiz + '/asp/relacao_master_.asp',qtd_62)
	oMenu.makeSub(posicao,5,'Atividade x Empresa',raiz + '/asp/relacao_ativ_emp_.asp',qtd_62)

	pos_Sub = 0
    qtd_62=3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cad_modulo.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar...................',raiz + '/asp/altera_modulo.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Excluir...................',raiz + '/asp/exclusao.asp?ID=1',qtd_62)

	pos_Sub = 1
    qtd_62=3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cad_atividade.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar...................',raiz + '/asp/altera_atividade.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Excluir...................',raiz + '/asp/exclusao.asp?ID=2',qtd_62)

	pos_Sub = 2
    qtd_62=3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cad_transacao.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar...................',raiz + '/asp/altera_transacao.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Excluir...................',raiz + '/asp/exclusao.asp?ID=3',qtd_62)

	pos_Sub = 3
    qtd_62=3
	oMenu.makeSubSub(posicao,pos_Sub,0,'Novo......................',raiz + '/asp/cad_empresa.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Alterar...................',raiz + '/asp/altera_empresa.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Excluir...................',raiz + '/asp/exclusao.asp?ID=4',qtd_62)
}

function funcao_IndexF(posicao)
{	//*** IndexF.js

	oMenu.makeMain(posicao,'FUNÇÃO',0)

	qtd_3=10;
    //oMenu.makeSub(posicao,0,'Função',0,qtd_3)
    oMenu.makeSub(posicao,0,'Função x Trasações',0,qtd_3)
	oMenu.makeSub(posicao,1,'Ag.Ativ x Ativ x Tran',raiz + '/asp/relacao_master_.asp',qtd_3)
	oMenu.makeSub(posicao,2,'Função Conflitante',raiz + '/asp/funcao/func_confl.asp?',qtd_3)
    oMenu.makeSub(posicao,3,'Função x Assunto',0,qtd_3)
    //oMenu.makeSub(posicao,5,'Orient Gerais',0,qtd_3)
    //oMenu.makeSub(posicao,6,'Termos Gerais',0,qtd_3)
    oMenu.makeSub(posicao,4,'Orient Mega',0,qtd_3)
    oMenu.makeSub(posicao,5,'Termos Mega ',0,qtd_3)
    oMenu.makeSub(posicao,6,'Assunto Mega ',0,qtd_3)
	oMenu.makeSub(posicao,7,'Orie.Funcao(OBS)',raiz + '/asp/orie_mape/seleciona_funcao.asp?pOpt=CF',qtd_3)
	oMenu.makeSub(posicao,8,'Libera Mape-Func',raiz + '/asp/funcao/sel_func_libera_mapeamento.asp?pOpt=1',qtd_3)
	oMenu.makeSub(posicao,9,'Libera Mape-Perfil',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=4',qtd_3)
	
	//qtd_30=3;	
	//oMenu.makeSubSub(posicao,0,0,'Nova',raiz + '/asp/funcao/cad_funcao.asp',qtd_30)
	//oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=1',qtd_30)
	//oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=2',qtd_30)
	
	qtd_30=2;
	posicao2 = 0
	oMenu.makeSubSub(posicao,posicao2,0,'Fun-Tra (mesmo MEGA)',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=3',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,1,'Fun-Tra (outro MEGA)',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=4',qtd_30)
	
	qtd_30=2;	
	posicao2 = 3	
	oMenu.makeSubSub(posicao,posicao2,0,'Altera assunto em massa',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=6',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,1,'Rel Função sem Assunto',raiz + '/asp/funcao/inclui_ori_gerais_mapeamento.asp',qtd_30)

	//qtd_30=3;	
	//oMenu.makeSubSub(posicao,5,0,'Inclui Orientações',raiz + '/asp/orie_mape/inclui_ori_gerais_mapeamento.asp',qtd_30)
	//oMenu.makeSubSub(posicao,5,1,'Altera Orientações',raiz + '/asp/orie_mape/seleciona_ori_gerais_mapeamento.asp?pOpt=A',qtd_30)
	//oMenu.makeSubSub(posicao,5,2,'Excluir Orientações',raiz + '/asp/orie_mape/seleciona_ori_gerais_mapeamento.asp?pOpt=E',qtd_30)
	
	//qtd_30=3;	
	//oMenu.makeSubSub(posicao,6,0,'Inclui Termos',raiz + '/asp/orie_mape/inclui_ori_gerais_mape_termos.asp',qtd_30)
	//oMenu.makeSubSub(posicao,6,1,'Altera Termos',raiz + '/asp/orie_mape/seleciona_ori_gerais_mape_termos.asp?pOpt=A',qtd_30)
	//oMenu.makeSubSub(posicao,6,2,'Excluir Termos',raiz + '/asp/orie_mape/seleciona_ori_gerais_mape_termos.asp?pOpt=E',qtd_30)

	qtd_30=3;	
	posicao2 = 4
	oMenu.makeSubSub(posicao,posicao2,0,'Inclui Orientações Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IO',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,1,'Altera Orientações Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_mapeamento.asp?pOpt=AO',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,2,'Excluir Orientações Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_mapeamento.asp?pOpt=EO',qtd_30)
	
	qtd_30=3;	
	posicao2 = 5
	oMenu.makeSubSub(posicao,posicao2,0,'Inclui Termos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IT',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,1,'Altera Termos Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_termo.asp?pOpt=AT',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,2,'Excluir Termos Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_termo.asp?pOpt=ET',qtd_30)

	qtd_30=3;	
	posicao2 = 	6
	oMenu.makeSubSub(posicao,posicao2,0,'Inclui Sub-Módulo Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,1,'Altera Sub-Módulo Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=AM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,posicao2,2,'Excluir Sub-Módulo Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=EM&pOpt2=M',qtd_30)
}

function perfil(posicao)
{	//*** IndexF.js

	oMenu.makeMain(posicao,'PERFIL',0)

	qtd_3=2;
    oMenu.makeSub(posicao,0,'Macro',0,qtd_3)
	//oMenu.makeSub(posicao,1,'Encam. Status',0,qtd_3)
    oMenu.makeSub(posicao,1,'Micro',0,qtd_3)
	//oMenu.makeSub(posicao,3,'Encam. Status',0,qtd_3)

	qtd_30=5;
	oMenu.makeSubSub(posicao,0,0,'Novo',raiz + '/asp/macroperfil/incluir_macro_perfil.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=1',qtd_30)
	oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=2',qtd_30)
	oMenu.makeSubSub(posicao,0,3,'Elaboração->Aprovação',raiz + '/asp/macroperfil/selec_valida_status1.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,4,'Validar Aprovação',raiz + '/asp/macroperfil/selec_valida_status2.asp',qtd_30)
	
	//qtd_33=1
	//oMenu.makeSubSub(posicao,1,0,'Elaboração->Aprovação',raiz + '/asp/macroperfil/selec_valida_status1.asp',qtd_33)

	qtd_30=4;
	oMenu.makeSubSub(posicao,1,0,'Novo',raiz + '/asp/microperfil/incluir_micro_perfil.asp?pOPT=1',qtd_30)
	oMenu.makeSubSub(posicao,1,1,'Alterar',raiz + '/asp/microperfil/seleciona_micro_perfil.asp?pOPT=2',qtd_30)
	oMenu.makeSubSub(posicao,1,2,'Excluir',raiz + '/asp/microperfil/seleciona_micro_perfil.asp?pOPT=3',qtd_30)
	oMenu.makeSubSub(posicao,1,3,'Elaboração->Criação',raiz + '/asp/microperfil/selec_valida_micro1.asp',qtd_30)
	
	//qtd_33=1
	//oMenu.makeSubSub(posicao,3,0,'Elaboração->Aprovação',raiz + '/asp/microperfil/selec_valida_micro1.asp',qtd_33)	
}

function perfil_IndexG(posicao)
{	//*** IndexG.js

	oMenu.makeMain(posicao,'PERFIL',0)

	qtd_3=2;
	oMenu.makeSub(posicao,0,'Criação R/3 Macro',0,qtd_3)
	oMenu.makeSub(posicao,1,'Criação R/3 Micro',0,qtd_3)

	qtd_33=1
	oMenu.makeSubSub(posicao,0,0,'Em Criação->Criado R/3',raiz + '/asp/macroperfil/selec_valida_status5.asp',qtd_33)
	
	qtd_33=1
	oMenu.makeSubSub(posicao,1,0,'Em Criação->Criado R/3',raiz + '/asp/microperfil/selec_valida_micro2.asp',qtd_33)
}

function perfil_IndexH(posicao)
{	//*** IndexH.js

	oMenu.makeMain(posicao,'PERFIL',0)

	qtd_3=2;
    oMenu.makeSub(posicao,0,'Macro',0,qtd_3)
	//oMenu.makeSub(posicao,1,'Encam. Status',0,qtd_3)
    oMenu.makeSub(posicao,1,'Micro',0,qtd_3)
	//oMenu.makeSub(posicao,3,'Encam. Status',0,qtd_3)

	qtd_30=6;
	oMenu.makeSubSub(posicao,0,0,'Novo',raiz + '/asp/macroperfil/incluir_macro_perfil.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=1',qtd_30)
	oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=2',qtd_30)
	oMenu.makeSubSub(posicao,0,3,'Elaboração->Aprovação',raiz + '/asp/macroperfil/selec_valida_status1.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,4,'Validar Aprovação',raiz + '/asp/macroperfil/selec_valida_status2.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,5,'Em Criação->Criado R/3',raiz + '/asp/macroperfil/selec_valida_status5.asp',qtd_30)
		
	//qtd_33=1
	//oMenu.makeSubSub(posicao,1,0,'Elaboração->Aprovação',raiz + '/asp/macroperfil/selec_valida_status1.asp',qtd_33)

	qtd_30=5;
	oMenu.makeSubSub(posicao,1,0,'Novo',raiz + '/asp/microperfil/incluir_micro_perfil.asp?pOPT=1',qtd_30)
	oMenu.makeSubSub(posicao,1,1,'Alterar',raiz + '/asp/microperfil/seleciona_micro_perfil.asp?pOPT=2',qtd_30)
	oMenu.makeSubSub(posicao,1,2,'Excluir',raiz + '/asp/microperfil/seleciona_micro_perfil.asp?pOPT=3',qtd_30)
	oMenu.makeSubSub(posicao,1,3,'Elaboração->Criação',raiz + '/asp/microperfil/selec_valida_micro1.asp',qtd_30)
	oMenu.makeSubSub(posicao,1,4,'Em Criação->Criado R/3',raiz + '/asp/microperfil/selec_valida_micro2.asp',qtd_30)
	
	//qtd_33=1
	//oMenu.makeSubSub(posicao,3,0,'Elaboração->Aprovação',raiz + '/asp/microperfil/selec_valida_micro1.asp',qtd_33)
}

function funcao_IndexP(posicao)
{
	oMenu.makeMain(posicao,'FUNÇÃO',0)

	qtd_3=1;
    oMenu.makeSub(posicao,0,'Função',0,qtd_3)

	qtd_30=3;
	oMenu.makeSubSub(posicao,0,0,'Função Conflitante',raiz + '/asp/funcao/func_confl.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,1,'Função Críticas',raiz + '/asp/funcao/sel_func_critico.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,2,'Libera Mapeamento',raiz + '/asp/funcao/sel_func_libera_mapeamento.asp?pOpt=1',qtd_30)
}

function perfil_IndexP(posicao)
{	//*** IndexP.js

	oMenu.makeMain(posicao,'PERFIL',0)
	
	qtd_3=1;
    oMenu.makeSub(posicao,0,'Micro Perfil',0,qtd_3)

	qtd_30=1;
	oMenu.makeSubSub(posicao,0,0,'Libera Mapeamento',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=4',qtd_30)
}

function escopo_IndexQ(posicao)
{
	oMenu.makeMain(posicao,'D-ESCOPO',0)

	qtd_0=23;
	oMenu.makeSub(posicao,0,'Agrup.(Mstr List R3)',0,qtd_0)
	oMenu.makeSub(posicao,1,'Atividade',0,qtd_0)
	oMenu.makeSub(posicao,2,'Transação',0,qtd_0)
	oMenu.makeSub(posicao,3,'Empresa/Unidade',0,qtd_0)

	oMenu.makeSub(posicao,4,'Ag.Ativ x Ativ x Tran',raiz + '/asp/relacao_master_.asp',qtd_0)
	oMenu.makeSub(posicao,5,'Atividade x Empresa',raiz + '/asp/relacao_ativ_emp_.asp',qtd_0)    
	oMenu.makeSub(posicao,6,'Fale Conosco',raiz + '/asp/fale_conosco.asp',qtd_0)    
	oMenu.makeSub(posicao,7,'Cadastro de Sub-Modulos',raiz + '/asp/cad_submodulo.asp',qtd_0)
   	oMenu.makeSub(posicao,8,'Fechamento de Escopo',0,qtd_0)
	oMenu.makeSub(posicao,9,'Troca Servidor DB',raiz + '/troca_servidor.asp',qtd_0)
	oMenu.makeSub(posicao,10,'Cadastro de Evento',0,qtd_0)
	oMenu.makeSub(posicao,11,'Solicitação de Escopo','http://164.85.62.152/ComSinergia/DataBase33.nsf/frmChamado?OpenForm',qtd_0)
	oMenu.makeSub(posicao,12,'Consulta Solicitação','http://164.85.62.152/ComSinergia/DataBase33.nsf/1D503FF15BA7EBF683256D12004C6ECD?OpenPage',qtd_0)
	oMenu.makeSub(posicao,13,'Cadastro Dono', raiz + '/asp/cad_dono.asp',qtd_0)
	oMenu.makeSub(posicao,14,'Cons Trans em uma data', raiz + '/asp/consulta_transacao_decomp.asp?tipo=1',qtd_0)
	oMenu.makeSub(posicao,15,'Cons Trans entre datas', raiz + '/asp/consulta_transacao_decomp.asp?tipo=2',qtd_0)		
	oMenu.makeSub(posicao,16,'Importa_Usu_Treina', raiz + '/asp/golive/importa_usuario.asp?tipo=2',qtd_0)		
	oMenu.makeSub(posicao,17,'Importa_Usu_Mapeados', raiz + '/asp/golive/importa_usuario0.asp?tipo=2',qtd_0)		
	oMenu.makeSub(posicao,18,'Criação de Lote', raiz + '/asp/golive/seleciona_para_cria_lote.asp?tipo=1',qtd_0)		
	oMenu.makeSub(posicao,19,'Gera arq Saída', raiz + '/asp/golive/consulta_lote2.asp',qtd_0)		
	oMenu.makeSub(posicao,20,'Excluir Lote', raiz + '/asp/golive/consulta_lote_para_exclusao.asp',qtd_0)		
	oMenu.makeSub(posicao,21,'Trata Correlatos', raiz + '/asp/golive/trata_correlatos.asp',qtd_0)	
	oMenu.makeSub(posicao,22,'Rel Aprovados Função', raiz + '/asp/funcao/rel_aprovados_funcao.asp',qtd_0)	

	qtd_00=3;
	oMenu.makeSubSub(posicao,0,0,'Novo',raiz + '/asp/cad_modulo.asp',qtd_00)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/altera_modulo.asp',qtd_00)
	oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/exclusao.asp?ID=1',qtd_00)
    qtd_01=3;
	oMenu.makeSubSub(posicao,1,0,'Novo',raiz + '/asp/cad_atividade.asp',qtd_01)
	oMenu.makeSubSub(posicao,1,1,'Alterar',raiz + '/asp/altera_atividade.asp',qtd_01)
	oMenu.makeSubSub(posicao,1,2,'Excluir',raiz + '/asp/exclusao.asp?ID=2',qtd_01)
    qtd_02=3;
	oMenu.makeSubSub(posicao,2,0,'Novo',raiz + '/asp/cad_transacao.asp',qtd_02)
	oMenu.makeSubSub(posicao,2,1,'Alterar',raiz + '/asp/altera_transacao.asp',qtd_02)
	oMenu.makeSubSub(posicao,2,2,'Excluir',raiz + '/asp/exclusao.asp?ID=3',qtd_02)
    qtd_03=3;
	oMenu.makeSubSub(posicao,3,0,'Novo',raiz + '/asp/cad_empresa.asp',qtd_03)
	oMenu.makeSubSub(posicao,3,1,'Alterar',raiz + '/asp/altera_empresa.asp',qtd_03)
	oMenu.makeSubSub(posicao,3,2,'Excluir',raiz + '/asp/exclusao.asp?ID=4',qtd_03)
    qtd_04=3;
	oMenu.makeSubSub(posicao,8,0,'Fecha Escopo',raiz + '/asp/fecha_escopo.asp',qtd_04)
	oMenu.makeSubSub(posicao,8,1,'Consulta Escopo',raiz + '/asp/escopo/consulta_escopo.asp',qtd_04)
	oMenu.makeSubSub(posicao,8,2,'Excluir',raiz + '/asp/exclusao.asp?ID=4',qtd_04)
    qtd_04=3;
	oMenu.makeSubSub(posicao,10,0,'Novo',raiz + '/asp/escopo/incluir_evento.asp?ID=I',qtd_04)
	oMenu.makeSubSub(posicao,10,1,'Alterar',raiz + '/asp/escopo/Seleciona_evento.asp?ID=A',qtd_04)
	oMenu.makeSubSub(posicao,10,2,'Excluir',raiz + '/asp/escopo/Seleciona_evento.asp?ID=E',qtd_04)
}

function cenario_Composto(posicao)
{
	oMenu.makeMain(posicao,'CENÁRIO',0)

	qtd_2=2;
    oMenu.makeSub(posicao,0,'Cenário',0,qtd_2)	
    oMenu.makeSub(posicao,1,'Classe',0,qtd_2)

	qtd_20=15;
	oMenu.makeSubSub(posicao,0,0,'aaNovo',raiz + '/asp/cenario/cad_cenario.asp',qtd_20)
 	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/cenario/altera_cenario.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/cenario/excluir_cenario.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,3,'Copia Cenario Onda',raiz + '/asp/cenario/copia_cenario_onda.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,4,'Editar Cenário/Transação',raiz + '/asp/cenario/selec_Mega_Proc_Sub_Cenario2.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,5,'Editar Ordem Cenário',raiz + '/asp/cenario/sel_cenario_altera_sequencia.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,6,'Definir Fluxo Mestre',raiz + '/asp/cenario/sel_cenario_gera_fluxo.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,7,'Definir Fluxo Mestre de Ref.',raiz + '/asp/cenario/sel_cenario_refer_fluxo.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,8,'Alterar Escopo de Cenário', raiz + '/asp/cenario/altera_escopo.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,9,'Alterar Assunto em massa', raiz + '/asp/cenario/sel_cenario_altera_assunto.asp?txtOPT=1',qtd_20)
	oMenu.makeSubSub(posicao,0,10,'Rel Cenarios sem Assunto', raiz + '/asp/cenario/sel_cenario_altera_assunto.asp?txtOPT=2',qtd_20)
	oMenu.makeSubSub(posicao,0,11,'Consulta Escopo em uma data', raiz + '/asp/cenario/consulta_escopo.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,12,'Consulta Escopo entre datas', raiz + '/asp/cenario/consulta_escopo_entre_datas.asp',qtd_20)
	oMenu.makeSubSub(posicao,0,13,'Rel Cenarios sem Empresa', raiz + '/asp/cenario/sel_cenario_altera_assunto.asp?txtOPT=3',qtd_20)
	oMenu.makeSubSub(posicao,0,14,'Cadastro Pacote de Teste', raiz + '/asp/cenario/cad_hist_pacote.asp',qtd_20)

	qtd_21=1;
	oMenu.makeSubSub(posicao,1,0,'Alterar Classe',raiz + '/asp/cenario/alterar_classe.asp',qtd_21)
}

function funcao_IndexQ(posicao)
{
	oMenu.makeMain(posicao,'FUNÇÃO',0)

	qtd_3=14;
    oMenu.makeSub(posicao,0,'Função',0,qtd_3)
    oMenu.makeSub(posicao,1,'Função x Trasações',0,qtd_3)
	oMenu.makeSub(posicao,2,'Ag.Ativ x Ativ x Tran',raiz + '/asp/relacao_master_.asp',qtd_3)
	oMenu.makeSub(posicao,3,'Função Conflitante',raiz + '/asp/funcao/func_confl.asp?',qtd_3)
    oMenu.makeSub(posicao,4,'Função x Assunto',0,qtd_3)
    oMenu.makeSub(posicao,5,'Orient Gerais',0,qtd_3)
    oMenu.makeSub(posicao,6,'Termos Gerais',0,qtd_3)
    oMenu.makeSub(posicao,7,'Orient Mega',0,qtd_3)
    oMenu.makeSub(posicao,8,'Termos Mega ',0,qtd_3)
    oMenu.makeSub(posicao,9,'Assuntos Mega ',0,qtd_3)
	oMenu.makeSub(posicao,10,'Orie.Funcao(OBS)',raiz + '/asp/orie_mape/seleciona_funcao.asp?pOpt=CF',qtd_3)
	oMenu.makeSub(posicao,11,'Função Críticas',raiz + '/asp/funcao/sel_func_critico.asp',qtd_3)
	oMenu.makeSub(posicao,12,'Libera Mapeamento',raiz + '/asp/funcao/sel_func_libera_mapeamento.asp?pOpt=1',qtd_3)
	oMenu.makeSub(posicao,13,'Classif Funcao',raiz + '/asp/funcao/sel_func_libera_mapeamento.asp?pOpt=2',qtd_3)

	qtd_30=3;
	oMenu.makeSubSub(posicao,0,0,'Nova',raiz + '/asp/funcao/cad_funcao.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=1',qtd_30)
	oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=2',qtd_30)
	
	qtd_30=2;
	oMenu.makeSubSub(posicao,1,0,'Fun-Tra (mesmo MEGA)',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=3',qtd_30)
	oMenu.makeSubSub(posicao,1,1,'Fun-Tra (outro MEGA)',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=4',qtd_30)
	
	qtd_30=1;	
	oMenu.makeSubSub(posicao,4,0,'Altera assunto em massa',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=6',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,5,0,'Inclui Orientações',raiz + '/asp/orie_mape/inclui_ori_gerais_mapeamento.asp',qtd_30)
	oMenu.makeSubSub(posicao,5,1,'Altera Orientações',raiz + '/asp/orie_mape/seleciona_ori_gerais_mapeamento.asp?pOpt=A',qtd_30)
	oMenu.makeSubSub(posicao,5,2,'Excluir Orientações',raiz + '/asp/orie_mape/seleciona_ori_gerais_mapeamento.asp?pOpt=E',qtd_30)
	
	qtd_30=3;	
	oMenu.makeSubSub(posicao,6,0,'Inclui Termos',raiz + '/asp/orie_mape/inclui_ori_gerais_mape_termos.asp',qtd_30)
	oMenu.makeSubSub(posicao,6,1,'Altera Termos',raiz + '/asp/orie_mape/seleciona_ori_gerais_mape_termos.asp?pOpt=A',qtd_30)
	oMenu.makeSubSub(posicao,6,2,'Excluir Termos',raiz + '/asp/orie_mape/seleciona_ori_gerais_mape_termos.asp?pOpt=E',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,7,0,'Inclui Orientações Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IO',qtd_30)
	oMenu.makeSubSub(posicao,7,1,'Altera Orientações Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_mapeamento.asp?pOpt=AO',qtd_30)
	oMenu.makeSubSub(posicao,7,2,'Excluir Orientações Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_mapeamento.asp?pOpt=EO',qtd_30)
	
	qtd_30=3;	
	oMenu.makeSubSub(posicao,8,0,'Inclui Termos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IT',qtd_30)
	oMenu.makeSubSub(posicao,8,1,'Altera Termos Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_termo.asp?pOpt=AT',qtd_30)
	oMenu.makeSubSub(posicao,8,2,'Excluir Termos Mega',raiz + '/asp/orie_mape/seleciona_ori_mega_termo.asp?pOpt=ET',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,9,0,'Inclui Assuntos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=IM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,9,1,'Altera Assuntos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=AM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,9,2,'Excluir Assuntos Mega',raiz + '/asp/orie_mape/seleciona_mega_processo.asp?pOpt=EM&pOpt2=M',qtd_30)
}

function perfil_IndexQ(posicao)
{
	oMenu.makeMain(posicao,'PERFIL',0)

	qtd_3=13;
    oMenu.makeSub(posicao,0,'Macro',0,qtd_3)
	oMenu.makeSub(posicao,1,'Encam. Status',0,qtd_3)
    oMenu.makeSub(posicao,2,'Micro',0,qtd_3)
	oMenu.makeSub(posicao,3,'Encam. Status',0,qtd_3)
	oMenu.makeSub(posicao,4,'Libera Mapeamento',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=4',qtd_3)
    oMenu.makeSub(posicao,5,'Orient Gerais',0,qtd_3)
    oMenu.makeSub(posicao,6,'Termos Gerais',0,qtd_3)
    oMenu.makeSub(posicao,7,'Orient Mega',0,qtd_3)
    oMenu.makeSub(posicao,8,'Termos Mega ',0,qtd_3)
    oMenu.makeSub(posicao,9,'Assuntos Mega ',0,qtd_3)
	oMenu.makeSub(posicao,10,'Orie. Perfil',raiz + '/asp/orie_perfil/seleciona_funcao_macro_perfil.asp?pOpt=1',qtd_3)
	oMenu.makeSub(posicao,11,'Orientações Geral',raiz + '/asp/orie_perfil/relat_ori_gerais_mape_perfil.asp',qtd_3)
	oMenu.makeSub(posicao,12,'Orientações Mega',raiz + '/asp/orie_perfil/seleciona_funcao_macro_perfil.asp?pOpt=6',qtd_3)
	
	qtd_30=4;
	oMenu.makeSubSub(posicao,0,0,'Novo',raiz + '/asp/macroperfil/incluir_macro_perfil.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=1',qtd_30)
	oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=2',qtd_30)
	oMenu.makeSubSub(posicao,0,3,'Edita Objetos',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=3',qtd_30)
	
	qtd_33=3
	oMenu.makeSubSub(posicao,1,0,'Elaboração->Aprovação',raiz + '/asp/macroperfil/selec_valida_status1.asp',qtd_33)
	oMenu.makeSubSub(posicao,1,1,'Validar Aprovação',raiz + '/asp/macroperfil/selec_valida_status2.asp',qtd_33)
	oMenu.makeSubSub(posicao,1,2,'Em Criação->Criado R/3',raiz + '/asp/macroperfil/selec_valida_status5.asp',qtd_33)

	qtd_30=4;
	oMenu.makeSubSub(posicao,2,0,'Novo',raiz + '/asp/microperfil/incluir_micro_perfil.asp?pOPT=1',qtd_30)
	oMenu.makeSubSub(posicao,2,1,'Alterar',raiz + '/asp/microperfil/seleciona_micro_perfil.asp?pOPT=2',qtd_30)
	oMenu.makeSubSub(posicao,2,2,'Excluir',raiz + '/asp/microperfil/seleciona_micro_perfil.asp?pOPT=3',qtd_30)
	oMenu.makeSubSub(posicao,2,3,'Edita Objetos',raiz + '/asp/microperfil/seleciona_micro_perfil.asp?pOPT=4',qtd_30)
	
	qtd_33=2
	oMenu.makeSubSub(posicao,3,0,'Elaboração->Aprovação',raiz + '/asp/microperfil/selec_valida_micro1.asp',qtd_33)
	oMenu.makeSubSub(posicao,3,1,'Em Criação->Criado R/3',raiz + '/asp/microperfil/selec_valida_micro2.asp',qtd_33)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,5,0,'Inclui Orientações',raiz + '/asp/orie_perfil/inclui_ori_gerais_perfil.asp',qtd_30)
	oMenu.makeSubSub(posicao,5,1,'Altera Orientações',raiz + '/asp/orie_perfil/seleciona_ori_gerais_mape_perfil.asp?pOpt=A',qtd_30)
	oMenu.makeSubSub(posicao,5,2,'Excluir Orientações',raiz + '/asp/orie_perfil/seleciona_ori_gerais_mape_perfil.asp?pOpt=E',qtd_30)
	
	qtd_30=3;	
	oMenu.makeSubSub(posicao,6,0,'Inclui Termos',raiz + '/asp/orie_perfil/inclui_ori_gerais_mape_perfil_termos.asp',qtd_30)
	oMenu.makeSubSub(posicao,6,1,'Altera Termos',raiz + '/asp/orie_perfil/seleciona_ori_gerais_mape_perfil_termos.asp?pOpt=A',qtd_30)
	oMenu.makeSubSub(posicao,6,2,'Excluir Termos',raiz + '/asp/orie_perfil/seleciona_ori_gerais_mape_perfil_termos.asp?pOpt=E',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,7,0,'Inclui Orientações Mega',raiz + '/asp/orie_perfil/seleciona_mega_processo.asp?pOpt=IO',qtd_30)
	oMenu.makeSubSub(posicao,7,1,'Altera Orientações Mega',raiz + '/asp/orie_perfil/seleciona_ori_mega_mape_perfil.asp?pOpt=AO',qtd_30)
	oMenu.makeSubSub(posicao,7,2,'Excluir Orientações Mega',raiz + '/asp/orie_perfil/seleciona_ori_mega_mape_perfil.asp?pOpt=EO',qtd_30)
	
	qtd_30=3;	
	oMenu.makeSubSub(posicao,8,0,'Inclui Termos Mega',raiz + '/asp/orie_perfil/seleciona_mega_processo.asp?pOpt=IT',qtd_30)
	oMenu.makeSubSub(posicao,8,1,'Altera Termos Mega',raiz + '/asp/orie_perfil/seleciona_ori_mega_perfil_termo.asp?pOpt=AT',qtd_30)
	oMenu.makeSubSub(posicao,8,2,'Excluir Termos Mega',raiz + '/asp/orie_perfil/seleciona_ori_mega_perfil_termo.asp?pOpt=ET',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,9,0,'Inclui Assuntos Mega',raiz + '/asp/orie_perfil/seleciona_mega_processo.asp?pOpt=IM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,9,1,'Altera Assuntos Mega',raiz + '/asp/orie_perfil/seleciona_mega_processo.asp?pOpt=AM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,9,2,'Excluir Assuntos Mega',raiz + '/asp/orie_perfil/seleciona_mega_processo.asp?pOpt=EM&pOpt2=M',qtd_30)
}
	
function desenho_IndexQ(posicao)
{
	oMenu.makeMain(posicao,'DESENHO',raiz + '/msg_desenho.htm')
}

function cursos_IndexQ(posicao)
{
	oMenu.makeMain(posicao,'CURSOS',0)

	qtd_3=1;
    oMenu.makeSub(posicao,0,'Cursos',0,qtd_3)

	qtd_30=9;
	oMenu.makeSubSub(posicao,0,0,'Novo',raiz + '/asp/curso/cad_curso.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/curso/seleciona_curso.asp?option=6',qtd_30)
	oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/curso/seleciona_curso.asp?option=5',qtd_30)
	oMenu.makeSubSub(posicao,0,3,'Curso x Função x Trans',raiz + '/asp/curso/seleciona_curso.asp?option=2',qtd_30)
	oMenu.makeSubSub(posicao,0,4,'Curso x Transação',raiz + '/asp/curso/seleciona_curso.asp?option=1',qtd_30)
	oMenu.makeSubSub(posicao,0,5,'Curso x Cenário',raiz + '/asp/curso/seleciona_curso.asp?option=3',qtd_30)
	oMenu.makeSubSub(posicao,0,6,'Curso x Correlato',raiz + '/asp/curso/seleciona_curso.asp?option=8',qtd_30)
	oMenu.makeSubSub(posicao,0,7,'Pré Requisito',raiz + '/asp/curso/seleciona_curso.asp?option=4',qtd_30)
	oMenu.makeSubSub(posicao,0,8,'Lib.Manual (LM)',raiz + '/asp/treinamento/seleciona_lm.asp',qtd_30)
}

function cases_IndexQ(posicao)
{
	oMenu.makeMain(posicao,'CASES',raiz + '/asp/cenario/rel_geral.asp')
}

function pep_IndexQ(posicao)
{
	oMenu.makeMain(posicao,'PEP',0)

	qtd_7=2;
	oMenu.makeSub(posicao,0,'Seleção',raiz + '/asp/xpep/asp/seleciona_plano.asp',qtd_7)
	oMenu.makeSub(posicao,1,'Lista Projeto',raiz + '/asp/xpep/asp/lista_projeto.asp',qtd_7)
}
	
function goLive(posicao)
{	
	oMenu.makeMain(posicao,'GOLIVE',0)
	
	qtd_3=1;
	oMenu.makeSub(posicao,0,'Gera Arquivo',raiz + '/asp/golive/consulta_lote2.asp',qtd_3)
}
	
function perfil_IndexV(posicao)
{
	oMenu.makeMain(posicao,'PERFIL',0)

	qtd_3=10;
    oMenu.makeSub(posicao,0,'Macro',0,qtd_3)
    oMenu.makeSub(posicao,1,'Micro',0,qtd_3)
    oMenu.makeSub(posicao,2,'Orient Gerais',0,qtd_3)
    oMenu.makeSub(posicao,3,'Termos Gerais',0,qtd_3)
    oMenu.makeSub(posicao,4,'Orient Mega',0,qtd_3)
    oMenu.makeSub(posicao,5,'Termos Mega ',0,qtd_3)
    oMenu.makeSub(posicao,6,'Assuntos Mega ',0,qtd_3)	
	oMenu.makeSub(posicao,7,'Orientações Perfil',raiz + '/asp/orie_perfil/seleciona_funcao_macro_perfil.asp?pOpt=1',qtd_3)
	oMenu.makeSub(posicao,8,'Orientações Geral',raiz + '/asp/orie_perfil/relat_ori_gerais_mape_perfil.asp',qtd_3)
	oMenu.makeSub(posicao,9,'Orientações Mega',raiz + '/asp/orie_perfil/seleciona_funcao_macro_perfil.asp?pOpt=6',qtd_3)

	qtd_30=5;
	oMenu.makeSubSub(posicao,0,0,'Novo',raiz + '/asp/macroperfil/incluir_macro_perfil.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,1,'Alterar',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=1',qtd_30)
	oMenu.makeSubSub(posicao,0,2,'Excluir',raiz + '/asp/macroperfil/seleciona_macro_perfil.asp?pOPT=2',qtd_30)
	oMenu.makeSubSub(posicao,0,3,'Elaboração->Aprovação',raiz + '/asp/macroperfil/selec_valida_status1.asp',qtd_30)
	oMenu.makeSubSub(posicao,0,4,'Validar Aprovação',raiz + '/asp/macroperfil/selec_valida_status2.asp',qtd_30)
	
	qtd_30=4;
	oMenu.makeSubSub(posicao,1,0,'Novo',raiz + '/asp/microperfil/incluir_micro_perfil.asp?pOPT=1',qtd_30)
	oMenu.makeSubSub(posicao,1,1,'Alterar',raiz + '/asp/microperfil/seleciona_micro_perfil.asp?pOPT=2',qtd_30)
	oMenu.makeSubSub(posicao,1,2,'Excluir',raiz + '/asp/microperfil/seleciona_micro_perfil.asp?pOPT=3',qtd_30)
	oMenu.makeSubSub(posicao,1,3,'Elaboração->Criação',raiz + '/asp/microperfil/selec_valida_micro1.asp',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,2,0,'Inclui Orientações',raiz + '/asp/orie_perfil/inclui_ori_gerais_perfil.asp',qtd_30)
	oMenu.makeSubSub(posicao,2,1,'Altera Orientações',raiz + '/asp/orie_perfil/seleciona_ori_gerais_mape_perfil.asp?pOpt=A',qtd_30)
	oMenu.makeSubSub(posicao,2,2,'Excluir Orientações',raiz + '/asp/orie_perfil/seleciona_ori_gerais_mape_perfil.asp?pOpt=E',qtd_30)
	
	qtd_30=3;	
	oMenu.makeSubSub(posicao,3,0,'Inclui Termos',raiz + '/asp/orie_perfil/inclui_ori_gerais_mape_perfil_termos.asp',qtd_30)
	oMenu.makeSubSub(posicao,3,1,'Altera Termos',raiz + '/asp/orie_perfil/seleciona_ori_gerais_mape_perfil_termos.asp?pOpt=A',qtd_30)
	oMenu.makeSubSub(posicao,3,2,'Excluir Termos',raiz + '/asp/orie_perfil/seleciona_ori_gerais_mape_perfil_termos.asp?pOpt=E',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,4,0,'Inclui Orientações Mega',raiz + '/asp/orie_perfil/seleciona_mega_processo.asp?pOpt=IO',qtd_30)
	oMenu.makeSubSub(posicao,4,1,'Altera Orientações Mega',raiz + '/asp/orie_perfil/seleciona_ori_mega_mape_perfil.asp?pOpt=AO',qtd_30)
	oMenu.makeSubSub(posicao,4,2,'Excluir Orientações Mega',raiz + '/asp/orie_perfil/seleciona_ori_mega_mape_perfil.asp?pOpt=EO',qtd_30)
	
	qtd_30=3;	
	oMenu.makeSubSub(posicao,5,0,'Inclui Termos Mega',raiz + '/asp/orie_perfil/seleciona_mega_processo.asp?pOpt=IT',qtd_30)
	oMenu.makeSubSub(posicao,5,1,'Altera Termos Mega',raiz + '/asp/orie_perfil/seleciona_ori_mega_perfil_termo.asp?pOpt=AT',qtd_30)
	oMenu.makeSubSub(posicao,5,2,'Excluir Termos Mega',raiz + '/asp/orie_perfil/seleciona_ori_mega_perfil_termo.asp?pOpt=ET',qtd_30)

	qtd_30=3;	
	oMenu.makeSubSub(posicao,6,0,'Inclui Assuntos Mega',raiz + '/asp/orie_perfil/seleciona_mega_processo.asp?pOpt=IM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,6,1,'Altera Assuntos Mega',raiz + '/asp/orie_perfil/seleciona_mega_processo.asp?pOpt=AM&pOpt2=M',qtd_30)
	oMenu.makeSubSub(posicao,6,2,'Excluir Assuntos Mega',raiz + '/asp/orie_perfil/seleciona_mega_processo.asp?pOpt=EM&pOpt2=M',qtd_30)
}
		
function goLive_IndexV(posicao)
{	
	oMenu.makeMain(posicao,'GOLIVE',0)

	qtd_3=5;
	oMenu.makeSub(posicao,0,'Importa_Usu_Treina', raiz + '/asp/golive/importa_usuario.asp?tipo=2',qtd_3)		
	oMenu.makeSub(posicao,1,'Importa_Usu_Mapeados', raiz + '/asp/golive/importa_usuario0.asp?tipo=2',qtd_3)		
	oMenu.makeSub(posicao,2,'Prepara Lote', raiz + '/asp/golive/seleciona_para_cria_lote.asp?tipo=1',qtd_3)
    oMenu.makeSub(posicao,3,'Gera Arquivo',raiz + '/asp/golive/consulta_lote2.asp',qtd_3)
	oMenu.makeSub(posicao,4,'Excluir Lote', raiz + '/asp/golive/consulta_lote_para_exclusao.asp',qtd_3)			
}	
	


function consulta(posicao)
{
	oMenu.makeMain(posicao,'CONSULTA',0)

    qtd_6=9;
	oMenu.makeSub(posicao,0,'Processo',0,qtd_6)
	oMenu.makeSub(posicao,1,'Cenário',0,qtd_6)
	oMenu.makeSub(posicao,2,'Função',0,qtd_6)	
	oMenu.makeSub(posicao,3,'Perfil',0,qtd_6)
	oMenu.makeSub(posicao,4,'Curso',0,qtd_6)
	oMenu.makeSub(posicao,5,'Usuário',raiz + '/asp/consulta_usuario.asp',qtd_6)
	oMenu.makeSub(posicao,6,'Trans_Duplicada',0,qtd_6)
	oMenu.makeSub(posicao,7,'Escopo',0,qtd_6)	
	oMenu.makeSub(posicao,8,'Case',0,qtd_6)	
		
	pos_Sub = 0
	qtd_60=18
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
	oMenu.makeSubSub(posicao,pos_Sub,13,'Transacao sem decomp',raiz + '/asp/auditoria/sel_agrup_ativ_tran.asp',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,14,'Transacao sem Função',raiz + '/asp/auditoria/sel_Mega.asp?pOpt=1',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,15,'Curso sem Função',raiz + '/asp/auditoria/sel_Mega_Onda.asp?pOpt=1',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,16,'Curso sem Transacao',raiz + '/asp/auditoria/sel_Mega_Onda.asp?pOpt=2',qtd_60)
	oMenu.makeSubSub(posicao,pos_Sub,17,'Func/Tran não Curso',raiz + '/asp/auditoria/sel_Mega_Assunto.asp?pOpt=1',qtd_60)

pos_Sub = 1
	qtd_61=17
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
	oMenu.makeSubSub(posicao,pos_Sub,13,'Cenarios sem Responsável', raiz + '/asp/cenario/rel_resp.asp',qtd_61)
	//oMenu.makeSubSub(posicao,pos_Sub,14,'Desenvolvimentos', raiz + '/asp/cenario/rel_cenario_desenv.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,14,'Desenvolv. em Atraso', raiz + '/asp/cenario/rel_desenv_atrasado.asp',qtd_61)
	oMenu.makeSubSub(posicao,pos_Sub,15,'Consulta Validação de Escopo ', raiz + '/asp/cenario/consulta_escopo_cenario.asp',qtd_61)	
	oMenu.makeSubSub(posicao,pos_Sub,16,'Cenário Impactado em Desenv', raiz + '/asp/cenario/rel_impacto_desenv.asp',qtd_61)	
	

	pos_Sub = 2
	qtd_62=12
	oMenu.makeSubSub(posicao,pos_Sub,0,'Fun-Mega-Trans',raiz + '/asp/funcao/seleciona_Mega3.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Func-Trans',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=5',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Relatório Função',raiz + '/asp/funcao/rel_geral_funcao.asp?pOPT=1',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,3,'Função (Coluna)',raiz + '/asp/funcao/rel_geral_funcao.asp?pOPT=2',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,4,'Funções Conflitantes',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=8',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,5,'Rel Transação X Função',raiz + '/asp/funcao/consulta_func_trans.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,6,'Rel Funçao x Transação',raiz + '/asp/funcao/rel_func_trans_sem_rep.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,7,'Rel Função x Curso',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=9',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,8,'Rel Função sem Assunto',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=7',qtd_62)		
	oMenu.makeSubSub(posicao,pos_Sub,9,'Orientações Geral',raiz + '/asp/orie_mape/relat_ori_gerais_mapeamento.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,10,'Orientações Mega',raiz + '/asp/orie_mape/seleciona_funcao.asp?pOpt=RM',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,11,'Rel Função Sem Curso',raiz + '/asp/funcao/seleciona_funcao.asp?pOPT=10',qtd_62)

	pos_Sub = 3
    qtd_62=6
	oMenu.makeSubSub(posicao,pos_Sub,0,'Macro',raiz + '/asp/macroperfil/rel_geral_macro.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Solicitação Micro',raiz + '/asp/mIcroperfil/rel_mIcro.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Micro_R3',raiz + '/asp/mIcroperfil/rel_mIcro_r3.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,3,'Consulta Lote', raiz + '/asp/golive/consulta_lote2.asp?pAcao=C',qtd_62)		
	oMenu.makeSubSub(posicao,pos_Sub,4,'Orientações Geral',raiz + '/asp/orie_perfil/relat_ori_gerais_mapeamento.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,5,'Orientações Mega',raiz + '/asp/orie_perfil/seleciona_funcao_macro_perfil.asp?pOpt=RM',qtd_62)

	pos_Sub = 4
	qtd_62=7
	oMenu.makeSubSub(posicao,pos_Sub,0,'Curso',raiz + '/asp/curso/relat_geral_curso.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,1,'Curso x Função',raiz + '/asp/curso/seleciona_curso_rel.asp?option=2',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,2,'Curso x Transação',raiz + '/asp/curso/seleciona_curso_rel.asp?option=1',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,3,'Curso x Correlato',raiz + '/asp/curso/seleciona_curso_rel.asp?option=5',qtd_62)
	//oMenu.makeSubSub(posicao,pos_Sub,3,'Curso x Cenário',raiz + '/asp/curso/seleciona_curso_rel.asp?option=3',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,4,'Pré Requisito',raiz + '/asp/curso/seleciona_curso_rel.asp?option=4',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,5,'Rel.Compl.Cursos',raiz + '/asp/curso/curso_prerequisito.asp',qtd_62)
	oMenu.makeSubSub(posicao,pos_Sub,6,'Rel.Catálogo Cursos',raiz + '/asp/curso/sel_catalogo_curso.asp',qtd_62)
	
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
	
	pos_Sub = 8
    qtd_62=1
	oMenu.makeSubSub(posicao,pos_Sub,0,'Rel Case Condição Transação',raiz + '/asp/case/sel_case_condicao_transacao.asp',qtd_62)	

}	
//////////////////////////////////   FIM   ///////////////////////////////////////////////////////////////////////