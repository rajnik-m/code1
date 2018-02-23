
function DoNothing()
{
}

function SendPagetoPrinter() 
{
    var vBrowser=navigator.appName;
    if ((vBrowser=="Microsoft Internet Explorer"))
        {
            document.execCommand('print', false, null); 
        }
    else
        {
            window.print();
        }
}

function UpperCaseField(pText)
{
  pText.value = pText.value.toUpperCase()
}

function MaskEdit( pTextBox, pPositions, pDelimiter)
{
    var vValue = pTextBox.value
    var vPositions = pPositions.split(',')

    for (var i = 0; i <= vPositions.length; i++)
    {
	    for (var j = 0; j <= vValue.length; j++)
	    {
	        if (j == vPositions[i])
	        {
	            if (vValue.substring(j, j + 1) != pDelimiter)
	            {
	                vValue = vValue.substring(0, j) + pDelimiter + vValue.substring(j, vValue.length)
	            }
	        }
	    }
    }
    pTextBox.value = vValue
}

function CapitaliseWords(pText)
{
    var vWords = pText.value.split(/\s+/g) // split the sentence into an array of words
    for (var i = 0; i < vWords.length; i ++ )
    {
      var vFirstLetter = vWords[i].substring(0, 1).toUpperCase()
      var vRestOfWord = vWords[i].substring(1, vWords[i].length).toLowerCase()

      vWords[i] = vFirstLetter + vRestOfWord // re-assign it back to the array and move on
    }
    pText.value = vWords.join(' ') // join it back together
}

function PopupPicker(pDateControlName, pFindControl)
{
  var vCoordinates = GetObjectWindowPosition( pFindControl)
  var vValue = document.forms['nfpform'].elements[pDateControlName].value
  var vPopup=window.open('DatePicker.aspx?ID=' + pDateControlName + '&Value=' + vValue ,'DatePicker','width=240,height=200,top=' + vCoordinates.y + ',left=' + vCoordinates.x )
  vPopup.focus()
}

function SetDate(pValue)
{
    // retrieve the name of the input control on the parent form from the query string
    var vQueryParms = window.location.search.substr(1).substring(3)
    var vItems =  vQueryParms.split("&")
    // set the value of that control with the passed date
    window.opener.document.nfpform.elements[vItems[0]].value = pValue
    self.close()
}

// GetObjectWindowPosition(pObject)
// This function returns an object having .x and .y properties which are the coordinates of the parameter item
function GetObjectWindowPosition(pObject) 
{
	var vCoordinates = new Object()
  vCoordinates.x = FindPosX(pObject)
  vCoordinates.y = FindPosY(pObject)

	var x = 0
	var y = 0
	if (document.getElementById)
  {
		if (isNaN(window.screenX)) 
		{
		    if (document.documentElement)
		    {
			    x = vCoordinates.x - document.documentElement.scrollLeft + window.screenLeft
			    y = vCoordinates.y - document.documentElement.scrollTop + window.screenTop		        
		    }
		    else
		    {
			    x = vCoordinates.x - document.body.scrollLeft + window.screenLeft
			    y = vCoordinates.y - document.body.scrollTop + window.screenTop
			}
		}
		else
		{
			x = vCoordinates.x + window.screenX + (window.outerWidth - window.innerWidth) - window.pageXOffset
			y = vCoordinates.y + window.screenY + (window.outerHeight - 24 - window.innerHeight) - window.pageYOffset
		}
	}
	else if (document.all) 
	{
		x = vCoordinates.x - document.body.scrollLeft + window.screenLeft
		y = vCoordinates.y - document.body.scrollTop + window.screenTop
	}
	else if (document.layers) 
	{
		x = vCoordinates.x + window.screenX + (window.outerWidth - window.innerWidth) - window.pageXOffset
		y = vCoordinates.y + window.screenY + (window.outerHeight - 24 - window.innerHeight) - window.pageYOffset
	}
	vCoordinates.x = x
	vCoordinates.y = y
	return vCoordinates
}

function FindPosX(pObj)
{
	var vCurLeft = 0
	if (pObj.offsetParent)
	{
		while (pObj.offsetParent)
		{
			vCurLeft += pObj.offsetLeft
			pObj = pObj.offsetParent
		}
	}
	else if (pObj.x)
		vCurLeft += pObj.x
	return vCurLeft
}

function FindPosY(pObj)
{
	var vCurTop = 0
	if (pObj.offsetParent)
	{
		while (pObj.offsetParent)
		{
			vCurTop += pObj.offsetTop
			pObj = pObj.offsetParent
		}
	}
	else if (pObj.y)
		vCurTop += pObj.y
	return vCurTop
}
