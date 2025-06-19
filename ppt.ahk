; code 에서 저장시 스크립트 바로 적용
#HotIf WinActive("ppt.ahk - Visual Studio Code") 
^s:: {
    send "^s"
    RELOAD
}

#HotIf WinActive("ahk_class PPTFrameClass")
setPPT(){
    if not IsSet(oppt) {
        global oppt := 
        ComObjActive("PowerPoint.Application")
    }
    if not IsSet(sel) {
        global sel := oppt.ActiveWindow.selection
    }
}

; 0 select none
; 1 select slide
; 2 select shapeRange
; 3 select textRange
get_activeWindow() => ComObjActive("PowerPoint.Application").ActiveWindow
get_selection() => get_activeWindow().selection

is_slide_active() => get_ActiveWindow().ActivePane.ViewType = 1
is_none() => is_slide_active() and get_selection().Type = 0
is_slide() => is_slide_active() and get_selection().Type = 1
is_shape() => is_slide_active() and get_selection().Type = 2
is_text() => is_slide_active() and  get_selection().Type = 3

is_singleShape() => is_slide_active() and get_shapeRange().count = 1
is_multiShape() => is_slide_active() and get_shapeRange().count > 1
is_multiShapeOver2() => is_slide_active() and get_shapeRange().count > 2

get_shapeRange() => get_selection().HasChildShapeRange ?  get_selection().ChildShapeRange : get_selection().shapeRange
get_textRange() => get_selection().TextRange
get_textRange2() => get_selection().TextRange2

for_shape(func) {
    changed := false
    for s, _ in get_shapeRange() {
        changed := func(&s) or changed
    }
    return changed
}
for_shape2(func) {
    changed := false
    for s, _ in get_shapeRange(){
        if isset(begin) {
            changed := func(&begin, &s) or changed
        } else {
            begin := s
        }
    }
    return changed
}
do_nothing() {
}
mso(str) => ComObjActive("PowerPoint.Application").CommandBars.ExecuteMso(str)
if_(exp, func, hotkey) => exp() ? func() : send(hotkey)
if_2(exp1, func1, exp2, func2, hotkey) => exp1() ? func1() : exp2() ? func2() : send(hotkey)

a:: if_2(is_none, _1() => mso("ShapeRoundedRectangle"), is_multiShape, _2() => getBoundfor(genShapeBefore(5)), ThisHotkey)
d:: if_2(is_none, _1() => mso("ShapeOval"), is_multiShape, _2() => getBoundfor(genShapeBefore(9)), ThisHotkey)
t:: if_2(is_none, _1() => mso("TextBoxInsertHorizontal"), is_multiShape, _2() => getBoundfor(genTextBoxBefore()), ThisHotkey)
/:: if_2(is_none, _1() => mso("ShapeStraightConnectorArrow"), _2()=> is_multiShape() and for_shape2(addConnFirstOther), do_nothing, ThisHotkey)
^/:: if_2(is_none, _1() => mso("ShapeStraightConnectorArrow"), _2()=> is_multiShape() and for_shape2(addConnSeq), do_nothing, ThisHotkey)

selectionType(func_none ?, func_slide ?, func_shape ?, func_text ?) {
    setPPT()
    if (sel.Type = 0) and IsSet(func_none) {
        func_none()
    }else if (sel.Type = 1) and IsSet(func_slide){
        func_slide()
    }else if (sel.Type = 2) and IsSet(func_shape){
        func_shape()
    }else if (sel.Type = 3) and IsSet(func_text){
        func_text()
    }
}

multiShape3(afunc, bfunc){
    if is_multiShapeOver2(){
        return afunc
    }else if (IsSet(bfunc)) {
        return bfunc
    }
}

getCenterBoundfor(boxfunc){
    for s, _ in get_shapeRange(){
        if IsSet(left){
            left := s.left + s.width/2 < left ? s.left + s.width/2 : left
            right := right < s.left + s.width ? s.left + s.width/2 : right
            top := s.top + s.height/2 < top ? s.top + s.height/2 : top
            bottom := bottom < s.top + s.height/2 ? s.top + s.height/2 : bottom
        }else{
            left := s.left + s.width/2 , right:= left
            top := s.top + s.height/2 , bottom := top
        }
    }
    boxfunc(left, right, top, bottom)
}

centerAlign(left, right, top, bottom) => (right - left < bottom - top ) ?  mso("ObjectsAlignCenterHorizontalSmart") : mso("ObjectsAlignMiddleVerticalSmart")

getBoundfor(boxfunc){
    for s, _ in get_shapeRange(){
        if IsSet(left){
            left := s.left < left ? s.left : left
            right := right < s.left + s.width ? s.left + s.width : right
            top := s.top < top ? s.top : top
            bottom := bottom < s.top + s.height ? s.top + s.height : bottom
        }else{
            left := s.left , right:= left + s.width 
            top := s.top , bottom := top + s.height
        }
    }
    boxfunc(left, right, top, bottom)
}

genShapeBefore(num){
    genShape(left, right, top, bottom){
        ComObjActive("PowerPoint.Application").ActiveWindow.View.Slide.Shapes.AddShape(num, left-5, top-5, right - left +10, bottom - top + 10)
        mso("ObjectBringToFront")
    }
    return genshape
}
genTextBoxBefore(){
    genShape(left, right, top, bottom){
        ComObjActive("PowerPoint.Application").ActiveWindow.View.Slide.Shapes.AddTextBox(1, left-5, top-5, right - left +10, bottom - top + 10)
        mso("ObjectBringToFront")
    }
    return genshape
}
forShape(afunc, ufunc ?){
    multi() {
        changed := false
        if(sel.HasChildShapeRange){
            for s, _ in sel.ChildShapeRange{
                changed := afunc(&s) or changed
            } 
        }else{
            for s, _ in sel.shapeRange{
                changed := afunc(&s) or changed
            } 
        }
        if (! changed) and IsSet(ufunc){
            ufunc()
        }
    }
    return multi
}
addconn(&begin, &next){
    ncf := ComObjActive("PowerPoint.Application").ActiveWindow.View.Slide.Shapes.AddConnector(1,0,0,5,5)
    ncf.ConnectorFormat.BeginConnect ConnectedShape := begin, ConnectionSite := 1
    ncf.ConnectorFormat.EndConnect ConnectedShape := next, ConnectionSite := 1
    ncf.line.EndArrowHeadStyle := 2
    ncf.RerouteConnections
}
flow(&begin, &next, is_seq, func){
    if (begin.connector) {
        begin := next
    }else if (! next.connector){
        func(&begin, &next)
        if (is_seq)
            begin := next
        return true
    }
    return false
}
addConnFirstOther(&begin, &next) => flow(&begin,&next,False, addconn)
addConnSeq(&begin, &next) => flow(&begin,&next,True, addconn)
^Numpad2:: if_(is_multiShape, _1() => mso("ObjectsAlignBottomSmart"), "^2") ; 아래쪽 정렬
^Numpad4:: if_(is_multiShape, _1() => mso("ObjectsAlignLeftSmart"), "^4") ; 왼쪽 정렬
^Numpad5:: if_(is_multiShape, _1() => getCenterBoundfor(centerAlign), "^5") ; 중앙 정렬
^Numpad6:: if_(is_multiShape, _1() => mso("ObjectsAlignRightSmart"), "^6") ; 오른쪽 정렬
^Numpad8:: if_(is_multiShape, _1() => mso("ObjectsAlignTopSmart"), "^8") ; 위쪽 정렬

beginArrowToggle(&s) => s.connector ?  s.line.BeginArrowHeadStyle := Mod(s.line.BeginArrowHeadStyle , 2) + 1 : do_nothing()
endArrowToggle(&s) => s.connector ?  s.line.EndArrowHeadStyle := Mod(s.line.EndArrowHeadStyle , 2) + 1 : do_nothing()
,:: if_(is_shape, _0 => for_Shape(beginArrowToggle), ThisHotkey) ; 시작 화살표 토글
.:: if_(is_shape, _0 => for_Shape(endArrowToggle), ThisHotkey) ; 끝 화살표 토글
wordWrapToggle(&s) => s.HasTextFrame ? s.TextFrame.WordWrap := s.TextFrame.WordWrap = False : do_nothing()

NumpadEnd:: if_2(is_shape, _1() => for_Shape(wordWrapToggle), is_text, _2() => get_TextRange().WordWrap := get_TextRange().WordWrap = False, ThisHotkey) ; 도형의 텍스트 배치 토글 (도형이 글씨 영역을 제한할 때 사용)

shapeTextVerticalAlign(num){
    f(&s){
        if (s.hasTextFrame){
            s.TextFrame.VerticalAnchor := num
        }
    }
    return f
}
+Numpad2::{ ; 텍스트 아래 정렬
    selectionType(,,forShape(shapeTextVerticalAlign(4),),_3() => sel.ShapeRange.TextFrame.VerticalAnchor := 4)
}
+Numpad5::{ ; 텍스트 중앙 정렬
    selectionType(,,forShape(shapeTextVerticalAlign(3),),_3() => sel.ShapeRange.TextFrame.VerticalAnchor := 3)
}
+Numpad8::{ ; 텍스트 위쪽 정렬
    selectionType(,,forShape(shapeTextVerticalAlign(1),),_3() => sel.ShapeRange.TextFrame.VerticalAnchor := 1)
}
connChange(&s){
    if (s.connector) {
        s.connectorFormat.Type := Mod(s.connectorFormat.Type , 3) + 1
        if (s.connectorFormat.BeginConnected and s.connectorFormat.EndConnected) {
            s.RerouteConnections
        }
    }
    return s.connector
}
fillToggle(&s){
    if(! s.connector){
        if (s.fill.visible) and s.line.visible {
            s.fill.visible := False
        } else if s.fill.visible {
            s.line.visible := True
            s.fill.visible := True
        } else{
            s.fill.visible := True
            s.line.visible := False
        }
    }
    return ! s.connector
}
^NumpadDiv::{ ; 선 모양 변경 / ㄱ S  
    selectionType(,,forShape(connChange,),)
}
^NumpadAdd:: {
    selectionType(,,multiShape3(_2() => mso("AlignDistributeVertically"),_22() => false),)
}
^NumpadSub:: {
    selectionType(,,multiShape3(_2() => mso("AlignDistributeHorizontally"),_22() => false),)
}
^NumpadMult::{
    selectionType(,,forShape(fillToggle,_22() => Send("*")),_3() => Send("*"))
}
dashToggle(&s){
    s.line.DashStyle := Mod(s.line.DashStyle + 3,6)
    return true
}
^\::{
    if_(is_shape, forShape(dashToggle),ThisHotkey)
}
lineWidth(up){
    f(&s){
        if (up) {
            s.line.weight := s.line.weight + 0.25
        } else {
            s.line.weight := s.line.weight > 0 ? s.line.weight - 0.25 : 0
        }
        return True
    }
    return f
}
wheelup::{
    if_(is_shape, forShape(lineWidth(True)),"{WheelUp}")
}
wheelDown::{
    if_(is_shape, forShape(lineWidth(False)),"{WheelDown}")
}

MButton:: send "!hsfe" ; 도형 스포이드
+MButton:: send "!hsoe" ; 도형 아웃라인 스포이드
^MButton:: send "!hfce" ; 텍스트 스포이드 
!o:: {
    if (sel.Type = 0) 
    {
        MsgBox ("none")
    }
    else if (sel.Type = 2)
    {
        if sel.HasChildShapeRange = true
        {
            msg := "shape " . sel.ShapeRange.Count . " " . sel.ChildShapeRange.Count
        }
        else 
        {
            msg :=  "shape " . sel.ShapeRange.Count
        }
        MsgBox ( msg)
    }
    else if (sel.Type = 3)
        MsgBox ( "text " . sel.TextRange.Count)
    else if (sel.Type = 1)
        MsgBox ("slide")
    else 
        MsgBox (sel.Type)
}
;if (sel.Type = 2)
;{
;    !s::
;    {
;        sel.ShapeRange.Shadow.Visible := sel.ShapeRange.Shadow.Visible = False
;        ;sel.ShapeRange.Shadow.Blur := 10
;        ;sel.ShapeRange.Shadow.OffsetX :=3 
;        ;sel.ShapeRange.Shadow.OffsetY :=3 
;    }
;}



#HotIf 
^r::{
    send "^l^c"
    ClipWait
    htmlfile := ComObject("htmlfile")
    htmlfile.write("<meta http-equiv='X-UA-Compatible' content='IE=edge'>")
    JS := htmlfile.parentwindow
    JS.eval("var dataVar = decodeURI('" . A_Clipboard . "')")
    data := JS.dataVar
    MsgBox "Copied URL: " . data
    run "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe " . data

}