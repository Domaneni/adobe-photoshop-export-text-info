app.displayDialogs = DialogModes.NO;

/* =================== Define parametres =================== */
var csvPath = "C:\\Temp\\";
var a = new File(csvPath+app.activeDocument.name.substr(0, app.activeDocument.name.length-4) +".csv");

/* =================== Call main function =================== */
main();
a.close();

/* ====================== Functions ======================= */
function getLayerVisibilityByIndex( idx ) {
           var ref = new ActionReference();
           ref.putProperty( charIDToTypeID("Prpr") , charIDToTypeID( "Vsbl" ));
           ref.putIndex( charIDToTypeID( "Lyr " ), idx );
           return executeActionGet(ref).getBoolean(charIDToTypeID( "Vsbl" ));;
};

function getAllLayersByIndex(){
      function getNumberLayers(){
         var ref = new ActionReference();
         ref.putProperty( charIDToTypeID("Prpr") , charIDToTypeID("NmbL") )
         ref.putEnumerated( charIDToTypeID("Dcmn"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt") );
         return executeActionGet(ref).getInteger(charIDToTypeID("NmbL"));
      };
      function hasBackground() {
         var ref = new ActionReference();
         ref.putProperty( charIDToTypeID("Prpr"), charIDToTypeID( "Bckg" ));
         ref.putEnumerated(charIDToTypeID( "Lyr " ),charIDToTypeID( "Ordn" ),charIDToTypeID( "Back" ))//bottom Layer/background
         var desc =  executeActionGet(ref);
         var res = desc.getBoolean(charIDToTypeID( "Bckg" ));
         return res   
      };
      function getLayerType(idx,prop) {       
         var ref = new ActionReference();
         ref.putIndex(charIDToTypeID( "Lyr " ), idx);
         var desc =  executeActionGet(ref);
         var type = desc.getEnumerationValue(prop);
         var res = typeIDToStringID(type);
         return res   
      };
      var cnt = getNumberLayers()+1;
      var res = new Array();
      if(hasBackground()){
         var i = 0;
      }else{
         var i = 1;
      };
      var prop =  stringIDToTypeID("layerSection");
      for(i;i<cnt;i++){
         var temp = getLayerType(i,prop);
      if(temp != "layerSectionEnds") res.push(i);
   };
   return res;
};

function makeActiveByIndex( idx, visible ){
   var desc = new ActionDescriptor();
      var ref = new ActionReference();
      ref.putIndex(charIDToTypeID( "Lyr " ), idx)
      desc.putReference( charIDToTypeID( "null" ), ref );      
      desc.putBoolean( charIDToTypeID( "MkVs" ), visible );
   executeAction( charIDToTypeID( "slct" ), desc, DialogModes.NO );
};

function main(){
   // create new file for saving text info
   a.open('w');    
   // column names
   a.writeln("\"Name\",\"Path\",\"Font\",\"Size\",\"Style\",\"Letter_Spacing\",\"Line_Height\",\"Color\",\"Horizontal_Scale\",\"Vertical_Scale\",\"Underline\"");
    
   var groups = getAllLayersByIndex();
   
   // Go through all layers
   for(var i = 0; i < groups.length ; i++) {
            var visible = getLayerVisibilityByIndex(groups[i]);
            makeActiveByIndex( groups[i], true );
            var layer = app.activeDocument.activeLayer;
            
            if(layer.kind == LayerKind.TEXT){
                getTextInfo(layer);
            }
            
            if (!visible) {
                makeInvisibleByIndex(groups[i]); 
            }
   }
}

/* Get path of layer trought folders */
function getParentPath(parent){
   if(app.activeDocument.name == parent.name){
      return "";
   } else {
      return  getParentPath(parent.parent) + "/" + parent.name;
   }
}

function makeInvisibleByIndex(index) {
   var idHd = charIDToTypeID( "Hd  " );
   var desc54 = new ActionDescriptor();
   var idnull = charIDToTypeID( "null" );
   var list2 = new ActionList();
   var ref49 = new ActionReference();
   var idLyr = charIDToTypeID( "Lyr " );
   ref49.putIndex( idLyr, index );
   list2.putReference( ref49 );
   desc54.putList( idnull, list2 );
   executeAction( idHd, desc54, DialogModes.NO );
}

function getTextInfo(layer) {
        var layerPath = getParentPath(layer.parent);
        var tI = layer.textItem;
        if(tI.contents != "") {
            var fontName = "";
            var font = null;
            var styles = new Array();
            if(tI.font.typename == null) {
                fontName = "" + tI.font;                
            } else {
                font = Application.fonts.getByName(tI.font);
                fontName = font.family;
                styles = font.style.split(" ");
            }
            var font = "";
            var name = layer.name;
            var resolution = app.activeDocument.resolution / 72;
            var size = tI.size.value * resolution;
            var underline = false;
            //var underline = layer.textItem.underline == "UnderlineType.UNDERLINEOFF" ? false : true;
            var style = " ";
            for(var i = 0; i < styles.length; i++) {
                style = style + styles[i] + " ";
            }
            style = style.substr(0, style.length-1);
            var letter_spacing =  (tI.tracking/1000);// + " em";
            var line_height = "";
            if(tI.useAutoLeading) {
                line_height = "auto";
            } else {
                line_height = tI.leading.value;
            }
            var rgbC = tI.color.rgb;
            var color = "("+ Math.round(rgbC.red) + "," + Math.round(rgbC.green) + ","+ Math.round(rgbC.blue) + ")";
            var horizontal_scale =  tI.horizontalScale;
            var vertical_scale =  tI.verticalScale;
            a.writeln("\""+name+"\","+"\""+layerPath+"\","+"\""+fontName+"\","+"\""+size+"\","+"\""+style+"\","+"\""+letter_spacing+"\","+"\""+line_height+"\","
                        +"\""+color+"\","+"\""+horizontal_scale+"\","+"\""+vertical_scale+"\","+"\""+underline+"\"");
        }
}
