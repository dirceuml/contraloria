function f_IsEmpty(s) {
    if (s == "" || s == null) return true; else return false;
}

function f_InvalidChars(s) {
    var InvalidChars = "&+='%#_";
    if (f_IsEmpty(s)) return false;
    for (var i = 0; i < s.length; i++)
        for (var j = 0; j < InvalidChars.length; j++)
            if (s.charAt(i) == InvalidChars.charAt(j)) {
                return true;
            }
    return false;
}

function f_Trim(s) {
    while (s.charAt(0) == " ") s = s.substring(1, s.length);
    if (s.length > 1)
        while (s.charAt(s.length - 1) == " ") s = s.substring(0, s.length - 1);
    return s;
}

function f_WordCount(s) {
    var cont = 0;
    for (var i = 0; i < s.length; i++)
        if (s.charAt(i) == " " && s.charAt(i + 1) != " ")
        cont++;
    cont = cont + 1;
    return cont;
}

function f_GetWord(s, n) {
    var ini = 0;
    var fin = 0;
    if (n == 1) {
        for (var i = 0; i < s.length; i++)
            if (s.charAt(i) == " ") {
            fin = i; break;
        }
    }
    else {
        var num = 1;
        while (num != n) {
            var FindWord = false;
            for (var i = ini; i < s.length; i++) {
                if (s.charAt(i) == " " && s.charAt(i + 1) != " ") {
                    ini = i + 1; num++; FindWord = true; break;
                }
            }
            if (FindWord == false) break;
        }
        for (var i = ini; i < s.length; i++) {
            if (s.charAt(i) == " ") {
                fin = i; break;
            }
        }
    }
    if (FindWord == false) return "";
    if (fin == 0) fin = s.length;
    return s.substring(ini, fin)
}

function f_CampoValido(objPalabrasId, cont) {
    var objPalabras = document.getElementById(objPalabrasId);
    palabra = objPalabras.value;
    palabra = f_Trim(palabra);
    if (f_InvalidChars(palabra)) {
        alert('Text search "' + palabra + '" contains invalid chars: "&+=\'%#_"');
        objPalabras.focus();
        return false;
    }
    //    if (f_WordCount(palabra) > 3) {
    //        alert('Text search "' + palabra + '" contains more of 3 words. Delete one');
    //        objPalabras.focus();
    //        return false;
    //    }
    //    var palabra1 = f_GetWord(palabra, 1)
    //    var palabra2 = f_GetWord(palabra, 2)
    //    var palabra3 = f_GetWord(palabra, 3)
    //    if (palabra1.length < 2 || (palabra2 != "" && palabra2.length < 2) || (palabra3 != "" && palabra3.length < 2)) {
    //        alert("Words contains less than 2 chars");
    //        objPalabras.focus();
    //        return false;
    //    }
    //    objPalabras.value = palabra1;
    //    if (palabra2 != "") objPalabras.value += " " + palabra2;
    //    if (palabra3 != "") objPalabras.value += " " + palabra3;
    var palabra1 = f_GetWord(palabra, 1)
    //alert(palabra1.length);
    if (palabra1.length < 3) {
            alert("Words contains less than 3 chars");
            objPalabras.focus();
            return false;
        }
    return true;
}

function f_PalabraReservada(s) {
    var sPalabra = s.toUpperCase();
    if (sPalabra == "INC") return true; 
    else return false;
}

function f_SelectAll(id) {
    var frm = document.forms[0];
    for (i = 0; i < frm.elements.length; i++) {
        if (frm.elements[i].type == "checkbox") {
            frm.elements[i].checked = document.getElementById(id).checked;
        }
    }
}

function btnAtras_onclick() {
    history.go(-1);
}












function MM_swapImgRestore() { //v3.0
    var i, x, a = document.MM_sr; for (i = 0; a && i < a.length && (x = a[i]) && x.oSrc; i++) x.src = x.oSrc;
}
function MM_findObj(n, d) { //v4.01
    var p, i, x; if (!d) d = document; if ((p = n.indexOf("?")) > 0 && parent.frames.length) {
        d = parent.frames[n.substring(p + 1)].document; n = n.substring(0, p);
    }
    if (!(x = d[n]) && d.all) x = d.all[n]; for (i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
    for (i = 0; !x && d.layers && i < d.layers.length; i++) x = MM_findObj(n, d.layers[i].document);
    if (!x && d.getElementById) x = d.getElementById(n); return x;
}
function MM_swapImage() { //v3.0
    var i, j = 0, x, a = MM_swapImage.arguments; document.MM_sr = new Array; for (i = 0; i < (a.length - 2); i += 3)
        if ((x = MM_findObj(a[i])) != null) { document.MM_sr[j++] = x; if (!x.oSrc) x.oSrc = x.src; x.src = a[i + 2]; }
}





function f_VerificaBusquedaComercial() {

    palabra = document.forms[0].txtDesComercialB.value;
    palabra = f_Trim(palabra);
    if (f_IsEmpty(palabra)) {
        alert("Se debe ingresar un texto");
        document.forms[0].txtDesComercialB.focus();
        return false;
    }
    if (f_InvalidChars(palabra)) {
        alert("El texto ingresado contiene caracteres invalidos: ' %");
        document.forms[0].txtDesComercialB.focus();
        return false;
    }
    if (f_WordCount(palabra) > 3) {
        alert("El texto ingresado contiene mas de 3 palabras");
        document.forms[0].txtDesComercialB.focus();
        return false;
    }
    var palabra1 = f_GetWord(palabra, 1)
    var palabra2 = f_GetWord(palabra, 2)
    var palabra3 = f_GetWord(palabra, 3)
    if (palabra1.length < 3 || (palabra2 != "" && palabra2.length < 3) || (palabra3 != "" && palabra3.length < 3)) {
        alert("La longitud de la(s) palabras debe ser de 3 o mas caracteres");
        document.forms[0].txtDesComercialB.focus();
        return false;
    }
    document.forms[0].palabraB1.value = palabra1
    document.forms[0].palabraB2.value = palabra2
    document.forms[0].palabraB3.value = palabra3
    return true;
}



function f_ValidaRUC(objRUCId) {
    var objRUC = document.getElementById(objRUCId)
    var RUC = objRUC.value;

    RUC = f_Trim(RUC);

    if (f_IsEmpty(RUC)) {
        alert("Please input a Tax ID");
        objRUC.focus();
        return false;
    }
    if (RUC.length < 8) {
        alert("Tax ID contains less than 8 characters");
        objRUC.focus();
        return false;
    }
    return true;
}