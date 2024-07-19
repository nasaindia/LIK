var workbook;
var sheetName;
var sheet;
var numRows;


function readFile()
{
    // console.log("indise readFile");
    const file = 'LIK_DATABASE.xlsx';
    fetch(file)
        .then(response => response.arrayBuffer())
        .then(data => {
            workbook = XLSX.read(data, {type: 'array'});
            sheetName = workbook.SheetNames[0];
            sheet = workbook.Sheets[sheetName];
            // console.log(sheet['G2'].v);

            const range = XLSX.utils.decode_range(sheet['!ref']);
            numRows = range.e.r;
            console.log(numRows);

        })
        .catch(error => {
            console.error('Error reading the excel file', error);
        })
}



function calculate()
{

    addx = 100;
    addy = window.innerHeight;
    addy = 1.05*addy;
    var canvas = document.createElement('canvas');
    var img = document.getElementById("image2");
    canvas.width = addx + img.width;
    canvas.height = addy + img.height;
    var cont = canvas.getContext('2d');
    cont.drawImage(img, addx, addy, img.width, img.height);

    multiplyX = 9;
    multiplyY = 9;
    rows = Math.floor(img.height/multiplyY);
    columns = Math.floor(img.width/multiplyX);
    var posx, posy;


    for(var i=0; i<rows; i++)
    {
        for(var j=0; j<columns; j++)
        {
            var div = document.createElement("div");
            div.style.width = "6px";
            div.style.height = "6px";
            div.style.borderRadius = "3px";
            
            div.style.position = "absolute";
            posx = (addx + (multiplyX*i));
            posy = (addy + (multiplyY*j));
            div.style.left = posx + "px";
            div.style.top = posy + "px";
            
            var pixelData = cont.getImageData(posx, posy, 1, 1).data;
            if(pixelData[3] == 0)
            {
                div.style.background = "#dedede";
            }
            else
            {
                div.style.background = "transparent";
            }
            
            div.style.color = "black";
            div.style.fontSize = "5px";
            document.body.appendChild(div);
        }
    }

    for(var i=0; i<numRows; i++)
    {
        var div = document.createElement("div");
        div.style.width = "6px";
        div.style.height = "6px";
        div.style.borderRadius = "3px";
        div.style.position = "absolute";
        var cellY = parseInt(i) + 2;
        div.style.left = (addx + Math.floor((multiplyX * (sheet['I' + cellY].v)*columns)/100)) + "px";
        console.log((sheet['I' + cellY].v));
        div.style.top = (addy + Math.floor((multiplyY * (sheet['J' + cellY].v)*rows)/100))  + "px";
        console.log((sheet['J' + cellY].v));
        div.style.background = "red";
        div.style.color = "black";
        div.style.fontSize = "5px";
        var idName = i;
        div.className = "red";
        div.setAttribute('data-index', idName);
        div.onmouseover = Change;
        // div.onmouseout = Default;
        document.body.appendChild(div);
    }

    
    // document.getElementById("i1").addEventListener("mouseover", mouseover(3));
}
function Change()
{
    var cont = this.getAttribute('data-index');
    var cellY = parseInt(cont) + 2;
    // console.log(cellY);
    // var cellAddress = cellX + cellY;
    // cellAddress = 'G2';
    // console.log(sheet[cellAddress]);
    document.getElementById("data").innerHTML = cont;
    this.style.cursor="pointer";
    document.getElementById("glink").href = sheet['K' + cellY].v;
    document.getElementById("cityT").innerHTML = sheet['G' + cellY].v;
    document.getElementById("collegeT").innerHTML = sheet['D' + cellY].v;
    document.getElementById("yearName").innerHTML = sheet['B' + cellY].v;
    document.getElementById("building").innerHTML = sheet['F' + cellY].v;
    document.getElementById("glink").style.visibility = "visible";
    document.getElementById("city").style.visibility = "visible";
    document.getElementById("yearName").style.visibility = "visible";
    document.getElementById("college").style.visibility = "visible";
    document.getElementById("building").style.visibility = "visible";
    document.getElementById("begin").style.visibility = "hidden";
    document.getElementById("participant").style.visibility = "visible";
}

function Default()
{
    document.getElementById("data").innerHTML = "hover on red dots to see";
}


