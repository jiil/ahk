; code 에서 저장시 스크립트 바로 적용
#HotIf WinActive("ppt.ahk - Visual Studio Code") 
^s:: {
    send "^s"
    RELOAD
}

#HotIf WinActive("ahk_class PPTFrameClass")
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
is_connInclue() => is_shape() and for_shape(_(&s) => s.connector)
is_table() => is_singleShape() and get_shapeRange().hasTable

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
        get_ActiveWindow().View.Slide.Shapes.AddShape(num, left-5, top-5, right - left +10, bottom - top + 10)
        mso("ObjectBringToFront")
    }
    return genshape
}
genTextBoxBefore(){
    genShape(left, right, top, bottom){
        get_ActiveWindow().View.Slide.Shapes.AddTextBox(1, left-5, top-5, right - left +10, bottom - top + 10)
        mso("ObjectBringToFront")
    }
    return genshape
}
addconn(&begin, &next){
    ncf := get_ActiveWindow().View.Slide.Shapes.AddConnector(1,0,0,5,5)
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
,:: if_(is_connInclue, _0() => for_Shape(beginArrowToggle), ThisHotkey) ; 시작 화살표 토글
.:: if_(is_connInclue, _0() => for_Shape(endArrowToggle), ThisHotkey) ; 끝 화살표 토글
wordWrapToggle(&s) => s.HasTextFrame ? s.TextFrame.WordWrap := s.TextFrame.WordWrap = False : do_nothing()

NumpadEnd:: if_2(is_shape, _1() => for_Shape(wordWrapToggle), is_text, _2() => wordWrapToggle(&_ :=get_ShapeRange()), "1") ; 도형의 텍스트 배치 토글 (도형이 글씨 영역을 제한할 때 사용)

shapeTextVerticalAlign(&s, num)=> s.hasTextFrame ? s.TextFrame.VerticalAnchor := num : do_nothing()
shapeTextVerticalTop(&s) => s.hasTable ? mso("TableCellAlignTop") : shapeTextVerticalAlign(&s, 1) ; text align top
shapeTextVerticalBottom(&s) => s.hasTable ? mso("TableCellAlignBottom") : shapeTextVerticalAlign(&s, 4) ; text align bottom
shapeTextVerticalCenter(&s) => s.hasTable ? mso("TableCellAlignCenterVertically") : shapeTextVerticalAlign(&s, 3) ; text align center
NumpadDown:: if_2(is_shape, _1() => for_Shape(shapeTextVerticalBottom), is_text, _2() => shapeTextVerticalBottom(&_ := get_shapeRange()), "2") ; 텍스트 아래 정렬
NumpadClear:: if_2(is_shape, _1() => for_Shape(shapeTextVerticalCenter), is_text, _2() => shapeTextVerticalCenter(&_ := get_shapeRange()), "5") ; 텍스트 중앙 정렬
NumpadUp:: if_2(is_shape, _1() => for_Shape(shapeTextVerticalTop), is_text, _2() => shapeTextVerticalTop(&_ := get_shapeRange()), "8") ; 텍스트 위 정렬
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
            s.fill.visible := False
        } else if s.line.visible {
            s.fill.visible := True
            s.line.visible := False
        } else {
            s.fill.visible := True
            s.line.visible := True
        }
    }
    return ! s.connector
}
^NumpadDiv:: if_(is_connInclue, _1() => for_Shape(connChange), "/") ; 선 모양 변경 / ㄱ S  
^NumpadAdd:: if_2(is_multiShapeOver2, _1() => mso("AlignDistributeVertically"), is_table, _2()=> mso("TableRowsDistribute") , "+") ; 수직 같은 간격 배치
^NumpadSub:: if_2(is_multiShapeOver2, _1() => mso("AlignDistributeHorizontally"), is_table, _2()=> mso("TableColumnsDistribute") , "-") ; 수평 같은 간격 배치
^NumpadMult:: if_2(is_shape, _1() => for_Shape(fillToggle), is_text, _2()=> fillToggle(&s := get_shapeRange()), "*") ; 도형 채우기 토글

samesize(&begin, &next){
    next.height := begin.height
    next.width := begin.width
}
^NumpadEnter::if_(is_multiShape, _() => for_shape2(samesize),"{Enter}")

dashToggle(&s) => s.line.DashStyle := Mod(s.line.DashStyle + 3,6)
^\:: if_2(is_shape, _()=>for_Shape(dashToggle),is_text, _2()=> dashToggle(&s:= get_shapeRange()),ThisHotkey)
cellsBorder(&s, is_up) {
    r :=1
    Loop s.Table.Rows.Count {
        c := 1
        Loop s.Table.Columns.Count {
            if (s.Table.Cell(r,c).selected) {
                if not IsSet(rmin) {
                    rmin := r
                    rmax := r
                }
                if not IsSet(cmin) {
                    cmin := c
                    cmax := c
                }else {
                    rmax := r > rmax ? r : rmax
                    cmax := c > cmax ? c : cmax
                }
            }
            c++
        }
        r++
    }
    Loop rmax - rmin + 1 {
        leftCell := s.Table.Cell(rmin + A_Index - 1, cmin)
        rightCell := s.Table.Cell(rmin + A_Index - 1, cmax)
        if is_up {
            leftCell.Borders.item(2).weight := leftCell.Borders.item(2).weight + 1
            rightCell.Borders.item(4).weight := rightCell.Borders.item(4).weight + 1
        } else {
            leftCell.Borders.item(2).weight := leftCell.Borders.item(2).weight < 1 ? 0: leftCell.Borders.item(2).weight - 1
            rightCell.Borders.item(4).weight := rightCell.Borders.item(4).weight < 1 ? 0: rightCell.Borders.item(4).weight - 1
        }
    }
    loop cmax - cmin + 1 {
        topCell := s.Table.Cell(rmin, cmin + A_Index - 1)
        bottomCell := s.Table.Cell(rmax, cmin + A_Index - 1)
        if is_up {
            topCell.Borders.item(1).weight := topCell.Borders.item(1).weight + 1
            bottomCell.Borders.item(3).weight := bottomCell.Borders.item(3).weight + 1
        } else {
            topCell.Borders.item(1).weight := topCell.Borders.item(1).weight < 1 ? 0: topCell.Borders.item(1).weight - 1
            bottomCell.Borders.item(3).weight := bottomCell.Borders.item(3).weight < 1 ? 0: bottomCell.Borders.item(3).weight - 1
        }
    }
}
lineWidthUp(&s) => s.hasTable ? cellsBorder(&s, true) : s.line.visible ? s.line.weight := s.line.weight + 0.25 : s.line.visible := true
lineWidthDown(&s) => s.hasTable ? cellsBorder(&s, false) : s.line.visible ? s.line.weight := s.line.weight > 0 ? s.line.weight - 0.25 : 0 : s.line.visible := true
wheelup:: if_2(is_shape, _()=>for_Shape(lineWidthUp),is_text, _1()=>get_textRange().font.size := get_textRange().font.size + 1,"{WheelUp}")
wheelDown:: if_2(is_shape, _()=>for_Shape(lineWidthDown),is_text, _1()=>get_textRange().font.size := get_textRange().font.size - 1,"{WheelDown}")

MButton:: send "!hsfe" ; 도형 스포이드
+MButton:: send "!hsoe" ; 도형 아웃라인 스포이드
^MButton:: send "!hfce" ; 텍스트 스포이드 


!o:: {
    if (is_none()) {
        MsgBox ("none")
    } else if (is_shape()) {
        if get_selection().HasChildShapeRange = true {
            msg := "shape " . get_ShapeRange().Count . " " . get_selection().ChildShapeRange.Count
        } else {
            msg :=  "shape " . get_ShapeRange().Count
        }
        MsgBox ( msg)
    }
    else if (is_text()) {
        MsgBox ( "text " . get_TextRange().Count)
    }else if (is_slide()){
        MsgBox ("slide")
    }else {
        MsgBox (get_selection().Type)
    }
}