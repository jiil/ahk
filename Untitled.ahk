; code 에서 저장시 스크립트 바로 적용
#HotIf WinActive("Untitled.ahk - Visual Studio Code") 
^s:: {
    send "^s"
    RELOAD
}



#HotIf WinActive("ahk_class PPTFrameClass")
setPPT(){
    if not IsSet(oppt) {
        global oppt := ComObjActive("PowerPoint.Application")
    }
    if not IsSet(sel) {
        global sel := oppt.ActiveWindow.selection
    }
}

; 0 select none
; 1 select slide
; 2 select shapeRange
; 3 select textRange

is_none() {
    setPPT()
    return sel.Type = 0
}
is_slide() {
    setPPT()
    return sel.Type = 1
}
is_shape() {
    setPPT()
    return sel.Type = 2
}
is_text() {
    setPPT()
    return sel.Type = 3
}
is_singleShape() {
    setPPT()
    if is_shape() {
        return sel.HasChildShapeRange ? sel.ChildShapeRange.count = 1 : sel.shapeRange.count = 1
    }else{
        return False
    }
}
is_multiShapeOverCnt(cnt) {
    setPPT()
    if is_shape() {
        return sel.HasChildShapeRange ? sel.ChildShapeRange.count > cnt : sel.shapeRange.count > cnt
    }else{
        return False
    }
}
is_multiShape(){
    return is_multiShapeOverCnt(1)
}
is_multiShape2(){
    return is_multiShapeOverCnt(2)

}
is_multiShape3(){
    return is_multiShapeOverCnt(3)
}
for_shape(func) {
    setPPT()
    changed := false
    range := sel.HasChildShapeRange ? sel.ChildShapeRange : sel.shapeRange
    for s, _ in sel.ChildShapeRange {
        changed := func(&s) or changed
    }
    return changed
}
for_shape2(func) {
    setPPT()
    changed := false
    range := sel.HasChildShapeRange ? sel.ChildShapeRange : sel.shapeRange
    for s, _ in range {
        if isset(begin) {
            changed := func(&begin, &s) or changed
        } else {
            begin := s
        }
    }
    return changed
}
mso(str) {
    oppt.CommandBars.ExecuteMso(str)
}
if_(exp, func, hotkey) {
    exp() ? func() : send(hotkey)
}
a::{
    if is_none(){
        mso("ShapeRoundedRectangle")
    } else if is_multiShape() {
        getBoundfor(genShapeBefore(5))
    } else {
        send(ThisHotkey)
    }
}
d::{
    if is_none() {
        mso("ShapeOval")
    } else if is_multiShape() {
        getBoundfor(genShapeBefore(9))
    } else {
        send(ThisHotkey)
    }
}
t::{
    if is_none() {
        mso("TextBoxInsertHorizontal")
    } else if is_multiShape() {
        getBoundfor(genTextBoxBefore())
    } else {
        send(ThisHotkey)
    }
}
/::{
    if is_none() {
        mso("ShapeStraightConnectorArrow")
    }else if is_multiShape() and for_shape2(addConnFirstOther) {
    }else{
        send(ThisHotkey)
    }
}
^/::{
    if is_none() {
        mso("ShapeStraightConnectorArrow")
    }else if is_multiShape() and for_shape2(addConnSeq) {
    }else{
        send(ThisHotkey)
    }
}

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

multiShape2(afunc, bfunc){
    if (sel.shapeRange.count > 1) or ((sel.HasChildShapeRange)and(sel.ChildShapeRange.count > 1)){
        return afunc
    }else if(IsSet(bfunc)) {
        return bfunc
    }
}

multiShape3(afunc, bfunc){
    if (sel.shapeRange.count > 2) or ((sel.HasChildShapeRange)and(sel.ChildShapeRange.count > 2)){
        return afunc
    }else if (IsSet(bfunc)) {
        return bfunc
    }
}

getCenterBoundfor(boxfunc){
    if (sel.HasChildShapeRange){
        for s, _ in sel.ChildShapeRange{
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
    }else{
        for s, _ in sel.shapeRange {
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
    }
    boxfunc(left, right, top, bottom)
}

centerAlign(left, right, top, bottom) {
    (right - left < bottom - top ) ?  mso("ObjectsAlignCenterHorizontalSmart") : mso("ObjectsAlignMiddleVerticalSmart")
}

getBoundfor(boxfunc){
    range := sel.HasChildShapeRange ? sel.ChildShapeRange : sel.shapeRange
    for s, _ in range {
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
        oppt.ActiveWindow.View.Slide.Shapes.AddShape(num, left-5, top-5, right - left +10, bottom - top + 10)
        mso("ObjectBringToFront")
    }
    return genshape
}
genTextBoxBefore(){
    genShape(left, right, top, bottom){
        oppt.ActiveWindow.View.Slide.Shapes.AddTextBox(1, left-5, top-5, right - left +10, bottom - top + 10)
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
forShape2(afunc,ufunc ?){
    multi() {
        changed := False
        if (sel.HasChildShapeRange) {
            for s, _ in sel.ChildShapeRange{
                if isset(begin) {
                    changed := afunc(&begin, &s) or changed
                } else {
                    begin := s
                }
            }
        }else{
            for s, _ in sel.shapeRange{
                if isset(begin) {
                    changed := afunc(&begin, &s) or changed
                } else {
                    begin := s
                }
            }
        }
        if (! changed) and IsSet(ufunc){
            ufunc()
        }
    }
    return multi
}
addConn(&begin, &next, seq){
    if (begin.connector) {
        begin := next
    }else if (! next.connector){
        ncf := oppt.ActiveWindow.View.Slide.Shapes.AddConnector(1,0,0,5,5)
        ncf.ConnectorFormat.BeginConnect ConnectedShape := begin, ConnectionSite := 1
        ncf.ConnectorFormat.EndConnect ConnectedShape := next, ConnectionSite := 1
        ncf.line.EndArrowHeadStyle := 2
        ncf.RerouteConnections
        if (seq)
            begin := next
        return true
    }
}
addConnFirstOther(&begin, &next){
    return addConn(&begin,&next,False)
}

addConnSeq(&begin, &next){
    return addConn(&begin,&next,True)
}
^Numpad2:: { ; 아래쪽 정렬
    selectionType(,,multiShape2(_2() => mso("ObjectsAlignBottomSmart"), _22() => send("2")), _4() => send("2"))
}
^Numpad4:: { ; 왼쪽 정렬
    selectionType(,,multiShape2(_2() => mso("ObjectsAlignLeftSmart"), _22() => send("4")), _4() => send("4"))
}
^Numpad5:: { ; 중앙 정렬
    selectionType(,,multiShape2(_2() => getCenterBoundfor(centerAlign), _22() => send("5")), _4() => send("5"))
}
^Numpad6:: { ; 오른쪽 정렬
    selectionType(,,multiShape2(_2() => mso("ObjectsAlignRightSmart"), _22() => send("6")), _4() => send("6"))
}
^Numpad8:: { ; 위쪽 정렬
    selectionType(,,multiShape2(_2() => mso("ObjectsAlignTopSmart"), _22() => send("8")), _4() => send("8"))
}
connArrowToggle(is_begin){
    f(&s) {
        if (s.connector) {
            if (is_begin)
                s.line.BeginArrowHeadStyle := Mod(s.line.BeginArrowHeadStyle , 2) + 1
            else
                s.line.EndArrowHeadStyle := Mod(s.line.EndArrowHeadStyle , 2) + 1
        }
        return s.connector
    }
    return f
}
,::{ ; 시작 선 / 화살표 토글 
    selectionType( _0() => send(",") , _1() => send(",")
        ,forShape(connArrowToggle(True), _22() => send(",")), _3() => send(","))
}
.:: { ; 끝 선 / 화살표 토글
    selectionType( _0() => send(".") , _1() => send(".")
        ,forShape(connArrowToggle(False), _22() => send(".")), _3() => send("."))
}
wordWrapToggle(&s){
    if(s.HasTextFrame) {
        s.TextFrame.WordWrap := s.TextFrame.WordWrap = False
    }
    return s.HasTextFrame
}
+Numpad1::{ ; 도형의 텍스트 배치 토글 (도형이 글씨 영역을 제한할 때 사용)
    selectionType(,,forShape(wordWrapToggle,), _3() => sel.ShapeRange.TextFrame.WordWrap := sel.ShapeRange.TextFrame.WordWrap = False)
}
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
fillout(&s){
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
        s.fill.visible := s.fill.visible = False
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
    selectionType(,,forShape(fillout,_22() => Send("*")),_3() => Send("*"))
}
dashToggle(&s){
    s.line.DashStyle := Mod(s.line.DashStyle + 3,6)
    return true
}
^\::{
    if_(is_shape, forShape(dashToggle),ThisHotkey)
    ;selectionType(,,forShape(dashToggle,_22() => Send(".")),_3() => Send("."))
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
    ;selectionType(_0 => Send("{WheelUp}"), _1 => Send("{WheelUp}"), forShape(lineWidth(True), _22() => Send("{WheelUp}")), _3 => Send("{WheelUp}")
}
wheelDown::{
    if_(is_shape, forShape(lineWidth(False)),"{WheelDown}")
    ;selectionType(_0 => send("{WheelDown}"), _1 => send("{WheelDown}"), forShape(lineWidth(False),_22() => Send("{WheelDown}")), _3 => send("{WheelDown}"))
}

MButton:: send "!hsfe" ; 도형 스포이드
+MButton:: send "!hsoe" ; 도형 아웃라인 스포이드
^MButton:: send "!hfce" ; 텍스트 스포이드 
!o:: {
if (sel.Type = 0) 
{
    MsgBox ("none" . oppt.ActivePresentation.slides.count)
    oppt.ActiveWindow.view.GoToSlide oppt.ActiveWindow.view.slide.slideIndex -1

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