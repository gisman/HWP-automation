# HWP-automation
python으로 HWP(아래한글) automation 구현의 일반적인 절차

## import

Windows Automation 사용을 위한 import

``` python
from xmlrpc.client import Boolean
import win32com.client as win32
```

## hwp 인스턴스 생성

-   필수: hwp 인스턴스를 생성
-   선택: 창을 보이지 않게 하고, 읽기 전용으로 설정.

``` python
hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible = False
hwp.EditMode = 0 # READ ONLY
```

### hwp 창 보이기

파일 변환 등의 자동화 작업을 할 때 창을 숨기는게 더 좋은 선택이지만,
디버그를 할 때는 hwp 창을 보여주는게 좋다.

``` python
hwp.XHwpWindows.Item(0).Visible = True # DEBUG
```

## 보안 승인

한글 컨트롤을 사용할 때 로컬 파일에 접근하거나 저장하려고 하면 보안 승인
메시지가 나타납니다.

보안 승인 메시지가 나타나지 않도록 처리하는 모듈을 설치하고,
레지스트리에 등록해야 함.

공식 문서: <https://developer.hancom.com/hwpctrl-hwpautomation/> 의 3.
보안 승인 모듈 참고:
<https://everyday-tech.tistory.com/entry/아래한글-자동화-보안모듈-등록>

``` python
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample") 
```

## hwp 유효성 검사

안 열리는 파일을 열려고 하면 영원히 응답하지 않는다. 따라서 열기 전에
hwp.GetFileInfo()로 유효한 파일인지 확인해야 한다.

GetFileInfo()의 반환값은 다음과 같다.

`Format  string  파일의 형식.`  
`    HWP : 한글 파일`  
`    UNKNOWN : 알 수 없음.`  
`VersionStr  string  파일의 버전 문자열  ex)5.0.0.3`  
`VersionNum  unsigned    int 파일의 버전 ex) 0x05000003`  
`Encrypted   int 암호 여부`

``` python
info=hwp.GetFileInfo(hwp_path)
if info.Item('Format') == 'HWP':
    # Do something
    # print(info.Item('Format'), info.Item('VersionStr'))
```

## hwp 파일 열기

공식 문서: <https://developer.hancom.com/hwpctrl-hwpautomation/> 의 58)
Open

-   suspendpassword: TRUE로 지정하면 암호가 있는 파일일 경우 암호를 묻지
    않고 무조건 읽기에 실패한 것

으로 처리한다.

-   forceopen: TRUE로 지정하면 읽기 전용으로 읽어야 하는 경우 대화상자를
    띄우지 않는다

``` python
b = hwp.Open(hwp_path, "", "suspendpassword:true;forceopen:true")
```

## Save As

현재 편집중인 문서를 지정한 이름으로 저장한다.

지원하는 포맷은 다양하다.

-   hwpx, hwp, owpml(개방형 표준 문서), docx, odt, html, xml,rtf, txt,
    csv, pdf, bmp, jpg, gif, png, wmf, emf

Save As pdf 예

``` python
if b:
    b = hwp.SaveAs(f"{res_dir}\\{pdf_file_name_strip}", "PDF", "")
```

## 4. hwp quit

파일을 닫고 한글을 종료한다.

``` python
hwp.Run("FileClose")
hwp.Quit()
