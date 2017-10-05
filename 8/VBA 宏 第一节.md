# VBA 宏 第一节

标签（空格分隔）： learn_excel

---
# 过程
在模块内的代码会被组织成过程，而过程会告诉应用程序如何去执行一个特定的任务。利用过程可将复杂的代码细分成许多部分，以便管理。

代码->过程=>执行的命令
# Sub过程
可以使用 Sub 过程去组织其它的过程，因此可以较容易了解并调试它们。在下面的示例中，Sub 过程 Main 传递参数值 56 去调用 Sub 过程 MultiBeep。运行 MultiBeep 后，控件返回 Main，然后 Main 调用 Sub 过程 Message。Message 显示一个信息框；当按“确定”键时，控件会返回 Main，接着 Main 退出执行。
```VBA
Sub Main()
    MultiBeep 56
    Message
End Sub

Sub MultiBeep(numbeeps)
    For counter = 1 To numbeeps
        Beep
    Next counter
End Sub

Sub Message()
    MsgBox "Time to take a break!"
End Sub
```
调用具有多个参数的 Sub 过程
下面的示例展示了调用具有多个参数的 Sub 过程的两种不同方法。当第二次调用 HouseCalc 时，因为使用 Call 语句所以需要利用括号将参数括起来。
```VBA
Sub Main()
    HouseCalc 99800, 43100
    Call HouseCalc(380950, 49500)
End Sub

Sub HouseCalc(price As Single, wage As Single)
    If 2.5 * wage <= 0.8 * price Then
        MsgBox "You cannot afford this house."
    Else
        MsgBox "This house is affordable."
    End If
End Sub
```
## Function过程
**在调用 Function 过程时使用括号**
为了使用函数的返回值，必须指定函数给变量，并且用括号将参数封闭起来；如下示例所示：
```VBA
Answer3 = MsgBox("Are you happy with your salary?", 4, "Question 3")
```
如果不在意函数的返回值，可以用调用 Sub 过程的方式来调用函数。如下面示例所示，可以省略括号，列出参数并且不要将函数指定给变量：
```VBA
MsgBox "Task Completed!", 0, "Task Box"
```

**小心** 
在上述例子中若包含括号，则语句会导致一个语法错误。

**传递命名参数**
Sub 或 Function 过程中的语句可以利用命名参数来传递值给被调用的过程。可以将命名参数以任何顺串行出。命名参数的组成是由参数名称紧接着冒号（:=）以及等号，然后指定一个值给参数。

下面的示例使用命名参数来调用不具返回值的 MsgBox 函数。
```VBA
MsgBox Title:="Task Box", Prompt:="Task Completed!"
```
下面的示例使用命名参数调用 MsgBox 函数。将返回值指定给变量 answer3。
```VBA
answer3 = MsgBox(Title:="Question 3", _
Prompt:="Are you happy with your salary?", Buttons:=4)
```

# 录制宏
开发工具
录制宏
使用相对引用（若不选择，则默认宏是绝对引用下的操作）
进行一系列操作
停止录制
再点宏->执行宏

#执行宏（单次操作）
点宏->执行宏
or
快捷键
or
插入-表单控件-按钮
or
快速访问工具栏->快捷键->宏

# 执行宏（一次性完成重复性操作）
开发工具-VB/ ALT+F11
模块存储了所有的VBA代码
![image_1brmh561eo9oh4d7ak1qr11u1c3h.png-10.6kB][1]
# VBA
## 对象
![image_1brme72aap6m1nq15gp14h9k102a.png-39.8kB][2]
![image_1brmdobr31bdj1u3c1f4u1il91h7f9.png-247.9kB][3]

多个同类型对象统称为**某个对象的集合**

引用对象
![image_1brmdp50c1mg44e118n91env11mgm.png-103.9kB][4]

## 属性

“帮助”里搜索“属性”的结果
![image_1brmdvn941kt6u7h5u829m1ftu1g.png-32.2kB][5]

## 方法 == 操作
对对象执行某个操作的动作
![image_1brme841a74oafcefhtsi3s2n.png-10.4kB][6]

方法= 对象.操作
```VBA
Worksheet.Add #增加一个新的工作簿
```
![image_1brme4cit1b011e7s1mg84ik50k1t.png-136.4kB][7]

##　属性与方法的区别
![image_1brmdsf2s3871bavv361jfjm8g13.png-5.8kB][8]
value. 后面的绿色文字都是方法，其余的都是属性

＃　编程环境　VBE 

VBE = VB Environment

## 打开VBE
通过开发工具的VB / ALT+F11 打开
右键工作表-查看代码
![image_1brmeeoad1j581qbnmetiph1je634.png-27.9kB][9]
 包括
对象
- sheet：存放工作表事件代码，仅在当前表调用
- ThisWorkBook：存放针对工作簿对象的代码
窗体：发挥工具箱的作用，存于窗体控件相关的代码，窗体代码没办法在其他窗口使用
模块：存放SUB过程和各种function过程的VBA代码
类模块：用户自定义，应用程序级别的事件

# 作业
![image_1brmi1fh5acqcc6fbkbboqjo3u.png-211.4kB][10]
```
Sub Macro1()
Dim i As Integer
For i = 1 To 53


'
' 加 Marco
'

'
    ActiveCell.Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-6]:RC[-1])"
    ActiveCell.Offset(1, 0).Range("A1").Select
    
    Next
    
End Sub

```


```
ActiveCell.Select #应该是对应“引用已被选中的Cell or 选择正在被激活的Cell”这个操作
ActiveCell.Offset(1, 0).Range("A1").Select  # 是什么意思不太明白
ActiveCell.FormulaR1C1 = "=SUM(RC[-6]:RC[-1])"  #FormulaR1C1的意思也尚未弄清楚,现在来猜是 对从当前位置左移6个cell至从当前位置左移1个cell的区间内的值进行加运算
```
# 操作分解
操作--选择K2
点击录制宏，双击K2这个单元格，然后结束录制

```
Sub select1()
'
' select1 Macro
'

'
    Range("K2").Select
End Sub
```

操作-plus
点击录制宏，单击K3单元格，直接输入“=”，选中H3，按回车，点击宏，选择PLUS进行编辑，最小化窗口，正常化窗口，点击K3并确认（查看）K3的公式，再次点K3，结束录制
```VBA
Sub PLUS()
'
' PLUS Macro
'

'
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    Range("K4").Select #按回车后自动跳到下一行，选中了K4
    Application.Goto Reference:="PLUS" #查看PLUS宏
    Application.WindowState = xlMinimized #窗口最小化
    Application.WindowState = xlNormal #窗口正常化
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    Range("K3").Select
End Sub

```


  [1]: http://static.zybuluo.com/419145138/qwaz2tdtjmpkbivvpoa2znlu/image_1brmh561eo9oh4d7ak1qr11u1c3h.png
  [2]: http://static.zybuluo.com/419145138/ruyh83d1jpvyqz9nsbcp194o/image_1brme72aap6m1nq15gp14h9k102a.png
  [3]: http://static.zybuluo.com/419145138/karc72nslfhsh43vy7fdim1a/image_1brmdobr31bdj1u3c1f4u1il91h7f9.png
  [4]: http://static.zybuluo.com/419145138/t9l924i0sfxfan71uhn7tnw3/image_1brmdp50c1mg44e118n91env11mgm.png
  [5]: http://static.zybuluo.com/419145138/ho24k6ygsdfckflb0fc9v14t/image_1brmdvn941kt6u7h5u829m1ftu1g.png
  [6]: http://static.zybuluo.com/419145138/08gnds6y2tzw9t0g50pip7db/image_1brme841a74oafcefhtsi3s2n.png
  [7]: http://static.zybuluo.com/419145138/vkwbo8cgkkek609c06liuzpt/image_1brme4cit1b011e7s1mg84ik50k1t.png
  [8]: http://static.zybuluo.com/419145138/pjicf5dern3r9bciqjnc5ifo/image_1brmdsf2s3871bavv361jfjm8g13.png
  [9]: http://static.zybuluo.com/419145138/id2r7jlom9eddek9klykixmj/image_1brmeeoad1j581qbnmetiph1je634.png
  [10]: http://static.zybuluo.com/419145138/pkr9sy8lkrfkj28dv96iu9pr/image_1brmi1fh5acqcc6fbkbboqjo3u.png