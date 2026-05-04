function getRandomID() {
    myRandomID = '';
    for (irandom=0;irandom<10;irandom++) {
        p = Math.floor(Math.random()*26);
        myRandomID += 'abcdefghijklmnopqrstuvwxyz'.substring(p,p+1);
    }
    return myRandomID;
}
function makeaqtree(ul,level) {
    var cn=ul.childNodes;
    var myReplacementNode = document.createElement("div");
    for (var icn=0;icn<cn.length;icn++) {
        if (cn[icn].nodeName.toUpperCase() != 'LI') {
            if (cn[icn].nodeName == '#text') {
                var isBlankNV = cn[icn].nodeValue.replace(/[\f\n\r\t\v ]*/,'');
                if (isBlankNV.length > 0) {
                    alert("UL structure is invalid; a UL contains a text node: '"+cn[icn].nodeValue+"'");
                    return;
                }
            } else {
                alert("UL structure is invalid; a UL contains something other than an LI (a "+cn[icn].nodeName+", in fact)");
                return;
            }
        }        
        var contentNodes = cn[icn].childNodes;
        var thereIsASubMenu = 0;
        var subNodes = new Array();
        for (var icontentNodes=0;icontentNodes<contentNodes.length;icontentNodes++) {
            var thisContentNode = contentNodes[icontentNodes];
            if (thisContentNode.nodeName == 'UL') {
                var subMenu = makeaqtree(thisContentNode,level+1);
                thereIsASubMenu = 1;
            } else {
                subNodes[subNodes.length] = thisContentNode.cloneNode(true);
            }
        }
        if (thereIsASubMenu) {
            var containerDiv = document.createElement("div");
            var containerElement = document.createElement("a");
            containerDiv.appendChild(containerElement);
            containerElement.className = "aqtree2link";
            
            var icon = document.createElement("span");
            icon.setAttribute("attachedsection",subMenu.getAttribute("id"));
            icon.setAttribute("href","#");
            icon.onclick = aqtree2ToggleVisibility;
            icon.innerHTML = aqtree2_expandMeHTML;
            icon.className = 'aqtree2icon';
            icon.id = 'icon-'+subMenu.id;
            containerElement.appendChild(icon);
        } else {
            var containerElement = document.createElement("div");
            var containerDiv = containerElement;
            if (subNodes.length > 0) {
                var icon = document.createElement("span");
                icon.innerHTML = aqtree2_bulletHTML;
                icon.className = 'aqtree2icon';
                containerElement.appendChild(icon);
            }
        }
        
        for (isubNodes=0;isubNodes<subNodes.length;isubNodes++) {
            sN = subNodes[isubNodes];
            if (sN.nodeName == '#text' && sN.nodeValue.replace(/[ \v\t\r\n]*/,'').length == 0) continue;
            containerElement.appendChild(sN);
        }
        if (thereIsASubMenu) {
            // now add the submenu itself!
            containerDiv.appendChild(subMenu);
        }
        myReplacementNode.appendChild(containerDiv);
    }
    var randID = getRandomID();
    myReplacementNode.setAttribute("id",randID);
    myReplacementNode.style.display = 'none';
    myReplacementNode.style.paddingLeft = (level*10)+'px';
    return myReplacementNode;
}
function makeaqtrees() {
    uls = document.getElementsByTagName("ul");
    for (iuls=0;iuls<uls.length;iuls++) {
        ULclassName = uls[iuls].className;
        if (ULclassName) {
            if (ULclassName.match(/\baqtree2\b/)) {
                returnNode = makeaqtree(uls[iuls],0);
                returnNode.style.display = 'block';
                pn = uls[iuls].parentNode;
                pn.replaceChild(returnNode,uls[iuls]);
            }        }    }}
function initaqtrees() {
    if (document.createElement &&
        document.getElementsByTagName &&
        RegExp &&
        document.body.innerHTML)
        makeaqtrees();
    else
        return;
}
window.onload = initaqtrees;
function aqtree2ToggleVisibility() {
    elemID = this.getAttribute("attachedsection");
    thisElem = document.getElementById(elemID);
    thisDisp = thisElem.style.display;
    thisElem.style.display = thisDisp == 'none' ? 'block' : 'none';
    icon = document.getElementById("icon-"+elemID);
    if (icon) icon.innerHTML = thisDisp == 'none' ? aqtree2_collapseMeHTML : aqtree2_expandMeHTML;
    return false;
}
aqtree2_expandMeHTML = '+';
aqtree2_collapseMeHTML = '-';
aqtree2_bulletHTML = '-';