

//Menux=new Array("texto que irá aparecer no menu","Link",nro de sub-elementos,altura,largura);
//see accompanying "config.htm" file for more information on structure of menus


/***********************************************************************************
*	(c) Ger Versluis 2000 version 5.41 24 December 2001	          *
*	For info write to menus@burmees.nl		          *
*	You may remove all comments for faster loading	          *		
***********************************************************************************/

	//var NoOffFirstLineMenus=2;			// Number of first level items
	var LowBgColor='#c0c0c0';			// Background color when mouse is not over
	var LowSubBgColor='#c0c0c0';			// Background color when mouse is not over on subs
	var HighBgColor='#195788';			// Background color when mouse is over
	var HighSubBgColor='#195788';			// Background color when mouse is over on subs
	var FontLowColor='#333333';			// Font color when mouse is not over
	var FontSubLowColor='black';			// Font color subs when mouse is not over
	var FontHighColor='#ffffff';			// Font color when mouse is over
	var FontSubHighColor='white';			// Font color subs when mouse is over
	var BorderColor='#909090';			// Border color
	var BorderSubColor='#909090';			// Border color for subs
	var BorderWidth=1;				// Border width
	var BorderBtwnElmnts=1;			// Border between elements 1 or 0
	var FontFamily="verdana, arial,comic sans ms,technical"	// Font family menu items
	var FontSize=7;				// Font size menu items
	var FontBold=1;				// Bold menu items 1 or 0
	var FontItalic=0;				// Italic menu items 1 or 0
	var MenuTextCentered='left';			// Item text position 'left', 'center' or 'right'
	var MenuCentered='left';			// Menu horizontal position 'left', 'center' or 'right'
	var MenuVerticalCentered='top';		// Menu vertical position 'top', 'middle','bottom' or static
	var ChildOverlap=.2;				// horizontal overlap child/ parent
	var ChildVerticalOverlap=.2;			// vertical overlap child/ parent
	var StartTop=117;				// Menu offset x coordinate
	var StartLeft=11;				// Menu offset y coordinate
	var VerCorrect=0;				// Multiple frames y correction
	var HorCorrect=0;				// Multiple frames x correction
	var LeftPaddng=3;				// Left padding
	var TopPaddng=0;				// Top padding
	var FirstLineHorizontal=1;			// SET TO 1 FOR HORIZONTAL MENU, 0 FOR VERTICAL
	var MenuFramesVertical=1;			// Frames in cols or rows 1 or 0
	var DissapearDelay=500;			// delay before menu folds in
	var TakeOverBgColor=1;			// Menu frame takes over background color subitem frame
	var FirstLineFrame='navig';			// Frame where first level appears
	var SecLineFrame='space';			// Frame where sub levels appear
	var DocTargetFrame='space';			// Frame where target documents appear
	var TargetLoc='';				// span id for relative positioning
	var HideTop=0;				// Hide first level when loading new document 1 or 0
	var MenuWrap=1;				// enables/ disables menu wrap 1 or 0
	var RightToLeft=0;				// enables/ disables right to left unfold 1 or 0
	var UnfoldsOnClick=0;			// Level 1 unfolds onclick/ onmouseover
	var WebMasterCheck=0;			// menu tree checking on or off 1 or 0
	var ShowArrow=1;				// Uses arrow gifs when 1
	var KeepHilite=1;				// Keep selected path highligthed
	var Arrws=['imagens/tri.gif',5,10,'',10,5,'imagens/trileft.gif',5,10];	// Arrow source, width and height

function BeforeStart(){return}
function AfterBuild(){return}
//function BeforeFirstOpen(){return}
//function AfterCloseAll(){return}


function BeforeFirstOpen(){
 if(ScLoc.HideArray){
  var H_A,H_Al,H_El,i;
  H_A=ScLoc.HideArray;
  H_Al=H_A.length;
  for (i=0;i<H_Al;i++){

 H_El=(Nav4)?ScLoc.document.layers[H_A[i]]:(DomYes)?ScLoc.document.getElementById(H_A[i]).style:ScLoc.document.all[H_A[i]].style;
   H_El.visibility=M_Hide}}}

function AfterCloseAll(){
 if(ScLoc.HideArray){
  var H_A,H_Al,H_El,i;
  H_A=ScLoc.HideArray;
  H_Al=H_A.length;
  for (i=0;i<H_Al;i++){

H_El=(Nav4)?ScLoc.document.layers[H_A[i]]:(DomYes)?ScLoc.document.getElementById(H_A[i]).style:ScLoc.document.all[H_A[i]].style;
   H_El.visibility=M_Show}}}




// Menu tree
//	MenuX=new Array(Text to show, Link, background image (optional), number of sub elements, height, width);
//	For rollover images set "Text to show" to:  "rollover:Image1.jpg:Image2.jpg"


