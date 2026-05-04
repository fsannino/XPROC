// Troca sem tirar da lista original e pode deletar a lista selecionada

sortitems = 1;

function move(fbox,tbox) 
	{
	for(var i=0; i<fbox.options.length; i++) 
		{
		if(fbox.options[i].selected && fbox.options[i].value != "") 
			{
			if(JaExiste(fbox.options[i].value,tbox))
				{
				var no = new Option();
				no.value = fbox.options[i].value;
				no.text = fbox.options[i].text;
				//fbox.options[i].selected = false
				tbox.options[tbox.options.length] = no;
				//tbox.options[tbox.options.length].selected = false
				//fbox.options[i].value = "";
				//fbox.options[i].text = "";
				}
			}
		}
	BumpUp(fbox);
	if (sortitems) SortD(tbox);
	}

function JaExiste(objvalor,tbox) 
	{
	for(var i=0; i<tbox.options.length; i++) 
		{
		//alert(tbox.options[i].value);
		//alert(objvalor);
		if(tbox.options[i].value == objvalor) 
			{
			alert("Este item já foi selecionado!")
			return false;
			}
		}
	return true;
	}

function BumpUp(box)  
	{
	for(var i=0; i<box.options.length; i++) 
		{
		if(box.options[i].value == "")  
			{
			for(var j=i; j<box.options.length-1; j++)  
				{
				box.options[j].value = box.options[j+1].value;
				box.options[j].text = box.options[j+1].text;
				}
			var ln = i;
			break;
			}
		}
	if(ln < box.options.length)  
		{
		box.options.length -= 1;
		BumpUp(box);
		}
	}

function SortD(box)  
	{
	var temp_opts = new Array();
	var temp = new Object();
	for(var i=0; i<box.options.length; i++)  
		{
		temp_opts[i] = box.options[i];
		}
	for(var x=0; x<temp_opts.length-1; x++)  
		{
		for(var y=(x+1); y<temp_opts.length; y++)  
			{
			if(temp_opts[x].text > temp_opts[y].text)  
				{
				temp = temp_opts[x].text;
				temp_opts[x].text = temp_opts[y].text;
				temp_opts[y].text = temp;
				temp = temp_opts[x].value;
				temp_opts[x].value = temp_opts[y].value;
				temp_opts[y].value = temp;
				}
			}
		}
	for(var i=0; i<box.options.length; i++)  
		{
		box.options[i].value = temp_opts[i].value;
		box.options[i].text = temp_opts[i].text;
		}
	}


function deleta(tbox) 
	// apaga os itens selecionados passando a lista
	{
	for(var i=0; i<tbox.options.length; i++) 
		{
		if(tbox.options[i].selected && tbox.options[i].value != "") 
			{
			tbox.options[i].value = "";
			tbox.options[i].text = "";			
			}
		}
	BumpUp(tbox);		
	if (sortitems) SortD(tbox);
	}
