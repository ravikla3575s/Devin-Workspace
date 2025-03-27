# MouseScroll.bas 詳細設計書

## 概要
MouseScroll.basはユーザーフォーム内でのマウススクロール機能を提供するモジュールです。スクロールフレーム内のコントロールをマウスホイールでスクロールするための機能を実装しています。

## 定数

### スクロール関連定数
```vba
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const MK_SHIFT As Long = &H4
Private Const MK_CONTROL As Long = &H8
Private Const WHEEL_DELTA As Long = 120
Private Const SB_LINEUP As Long = 0
Private Const SB_LINEDOWN As Long = 1
Private Const SB_PAGEUP As Long = 2
Private Const SB_PAGEDOWN As Long = 3
Private Const SB_THUMBPOSITION As Long = 4
Private Const SB_THUMBTRACK As Long = 5
Private Const SB_TOP As Long = 6
Private Const SB_BOTTOM As Long = 7
Private Const SB_ENDSCROLL As Long = 8
Private Const SB_HORZ As Long = 0
Private Const SB_VERT As Long = 1
Private Const SB_CTL As Long = 2
Private Const SB_BOTH As Long = 3
```
**説明**: Windows APIのスクロール関連メッセージと定数

## 主要機能

### InstallMouseWheelHandler
```vba
Public Sub InstallMouseWheelHandler(ByVal frm As Object, ByVal scrollFrame As Object)
```
**説明**: フォームにマウスホイールハンドラをインストールします。
**引数**: 
- `frm` (Object): マウスホイールハンドラをインストールするフォーム
- `scrollFrame` (Object): スクロール対象のフレーム
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `UserForm_Initialize` (DynamicShelfNameForm.frm)
- 呼び出し先: `SubclassForm`

### UninstallMouseWheelHandler
```vba
Public Sub UninstallMouseWheelHandler(ByVal frm As Object)
```
**説明**: フォームからマウスホイールハンドラをアンインストールします。
**引数**: 
- `frm` (Object): マウスホイールハンドラをアンインストールするフォーム
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `UserForm_Terminate` (DynamicShelfNameForm.frm)
- 呼び出し先: `UnsubclassForm`

### SubclassForm
```vba
Private Sub SubclassForm(ByVal frm As Object)
```
**説明**: フォームのウィンドウプロシージャをサブクラス化します。
**引数**: 
- `frm` (Object): サブクラス化するフォーム
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `InstallMouseWheelHandler`
- 呼び出し先: `SetWindowLong` (Windows API)

### UnsubclassForm
```vba
Private Sub UnsubclassForm(ByVal frm As Object)
```
**説明**: フォームのウィンドウプロシージャのサブクラス化を解除します。
**引数**: 
- `frm` (Object): サブクラス化を解除するフォーム
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `UninstallMouseWheelHandler`
- 呼び出し先: `SetWindowLong` (Windows API)

### WindowProc
```vba
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
```
**説明**: ウィンドウメッセージを処理するコールバック関数です。
**引数**: 
- `hwnd` (Long): ウィンドウハンドル
- `uMsg` (Long): メッセージID
- `wParam` (Long): メッセージの追加情報1
- `lParam` (Long): メッセージの追加情報2
**戻り値**: Long (メッセージ処理結果)
**呼び出し関係**:
- 呼び出し元: Windows OS
- 呼び出し先: `ProcessMouseWheel`, `CallWindowProc` (Windows API)

### ProcessMouseWheel
```vba
Private Sub ProcessMouseWheel(ByVal wParam As Long, ByVal lParam As Long)
```
**説明**: マウスホイールメッセージを処理します。
**引数**: 
- `wParam` (Long): マウスホイールメッセージの追加情報1
- `lParam` (Long): マウスホイールメッセージの追加情報2
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `WindowProc`
- 呼び出し先: `ScrollFrameVertically`

### ScrollFrameVertically
```vba
Private Sub ScrollFrameVertically(ByVal scrollAmount As Long)
```
**説明**: フレームを垂直方向にスクロールします。
**引数**: 
- `scrollAmount` (Long): スクロール量
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessMouseWheel`
- 呼び出し先: なし

## Windows API宣言

### SetWindowLong
```vba
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
```
**説明**: ウィンドウの属性を設定します。

### CallWindowProc
```vba
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
```
**説明**: 元のウィンドウプロシージャを呼び出します。

### GetScrollPos
```vba
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
```
**説明**: スクロールバーの現在位置を取得します。

### SetScrollPos
```vba
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
```
**説明**: スクロールバーの位置を設定します。

### SendMessage
```vba
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
```
**説明**: ウィンドウにメッセージを送信します。

## アルゴリズム詳細

### マウスホイールハンドリングアルゴリズム
1. フォームのウィンドウプロシージャをサブクラス化
2. マウスホイールメッセージ（WM_MOUSEWHEEL）をキャプチャ
3. ホイールの回転方向と量を取得
4. スクロールフレームの垂直スクロール位置を更新
5. フレーム内のコントロールの位置を調整

### スクロール処理アルゴリズム
1. 現在のスクロール位置を取得
2. スクロール量に基づいて新しい位置を計算
3. スクロールバーの位置を更新
4. フレーム内のコントロールの位置を調整

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、スクロール処理中にエラーが発生した場合でも処理が継続されるよう設計されています。

## 依存関係
- Windows API: ウィンドウメッセージ処理とスクロール機能に使用
