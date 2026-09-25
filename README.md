# PowerPoint 演示侧边栏

一个浮在屏幕上的竖条侧边栏，用来在放映幻灯片时快速操作：**上一页 / 下一页 / 笔 / 橡皮 / 结束放映**。

```
┌──────────────┐
│      PPT侧栏 │⋮  ← 标题栏，整条可拖
├──────────────┤
│      ﹀       │
│    上一页     │  ← 蓝条
├──────────────┤
│      ﹀       │
│    下一页     │  ← 蓝条
├──────────────┤
│      ✎ ▣     │  ← ▣ 是色块，点它换笔色
│      笔       │
├──────────────┤
│      ▭       │
│     橡皮      │  ← 双击 = 清空本页笔迹
├──────────────┤
│      ✕       │
│    结束放映   │  ← 红条（会先弹确认框）
└──────────────┘
```

---

## 一、先看这个：它是什么 / 不是什么

**它是什么**

一个独立的 Windows 小程序（`PowerPointSidebar.exe`）。它通过 COM 后期绑定（`PowerPoint.Application`）操控你**已经打开的** PowerPoint，所以功能上就是一个侧边栏。

**它不是什么 —— 请务必知悉**

它**不是**传统的 *in-process COM 加载项*（也就是「文件 → 选项 → 加载项 → COM 加载项」里那种、随 PowerPoint 进程一起加载的插件）。

原因：做传统 COM 加载项需要下面任意一套东西，而这台机器上都没有：

| 方案 | 需要的东西 | 本机状态 |
|---|---|---|
| VSTO 加载项 | .NET Framework 4.x + VSTO Runtime + Office PIA | ❌ 全都没有 |
| 共享 COM 加载项 | Office PIA + `regasm` 注册 | ❌ 没有 PIA |
| Office.js 加载项 | 一个 https 网址 + 清单文件 | 需要额外搭服务 |

探测结果（都已实测确认）：

- Visual Studio 只装了 `18\Insiders`，**没装 Office 开发工具**
- `VSTO Runtime` 注册表项不存在
- GAC 和 Office 目录里都**没有** `Microsoft.Office.Interop.PowerPoint.dll`
- 只有 .NET SDK `10.0.203`，没有 .NET Framework 4.x 引用包

所以这里走的是「**独立进程 + COM 操控**」路线。对你来说，操作体验和侧边栏完全一样。

> 如果你之后想要真正随 PowerPoint 启动的 COM 加载项，需要先装 **Visual Studio 的「Office/SharePoint 开发」工作负载**（它会带上 VSTO Runtime 和 PIA），到时候可以在这个项目基础上改。

---

## 二、怎么用

### 启动与自动行为

双击 `启动侧栏.cmd`，或直接运行：

```
PowerPointSidebar\bin\Release\net10.0-windows\PowerPointSidebar.exe
```

启动后它**不会弹出来**，而是安静地待在系统托盘（右下角）。接下来的行为全自动：

| 事件 | 侧栏行为 |
|---|---|
| 在 PowerPoint 里按 **F5** 开始放映 | **自动弹出**，贴到屏幕边缘，并盖在全屏放映之上 |
| 放映结束（按 Esc / 放完） | **自动收回托盘** |
| 关闭 PowerPoint | 侧栏继续待命，下次放映仍会自动弹出 |

侧栏**不会替你启动 PowerPoint**——只在检测到 `POWERPNT` 进程时才挂载。

常用命令行参数：

| 参数 | 作用 |
|---|---|
| 无参数 | 缩进托盘，等放映开始 |
| `--show` | 立即显示侧栏（不想等到放映时用，方便调试） |
| `--selftest` | 无头自检，打印与 PowerPoint 的连接状态后退出 |
| `--kill` | 结束所有正在运行的侧栏进程 |
| `--help` | 显示帮助 |

### 五个按钮的生效条件

按钮只在**放映状态**下生效。普通编辑视图下点击会提示「当前没有幻灯片放映窗口」。

### 连不上时怎么排查

侧栏带一个无头自检模式，不开窗口、不启动 PowerPoint，只把连接状态打出来：

```
PowerPointSidebar.exe --selftest
```

正常输出（本机实测）：

```
=== PowerPoint 侧边栏 自检 ===

1) PowerPoint 进程 (POWERPNT) 数量 : 1
2) COM 挂载 PowerPoint.Application  : 成功
3) 是否正在放映幻灯片           : 否
4) PowerPoint 主窗口句柄         : 0x70C6C

=> 结论：已连上 PowerPoint，但当前没有在放映。
        请在 PowerPoint 里按 F5 开始放映，按钮才会生效。

=== 自检结束 ===
```

输出同时会写到 exe 同目录的 `selftest.log`，从资源管理器双击运行时可以看那个文件。

`PowerPointSidebar.exe --help` 可以看全部命令行参数。

### 拖动与贴边

默认贴在**屏幕右边缘**，尺寸按毫米计算（**宽 38 mm × 高 77 mm**，适配大屏触控），
并垂直居中。

| 操作 | 效果 |
|---|---|
| 按住顶部空白处拖动 | 移动侧栏 |
| 拖到屏幕**左/右边缘** 40px 内松开 | **自动吸附贴边** |
| 拖到屏幕中间松开 | 自由悬浮 |
| 拖朝向幻灯片一侧的 10 mm 细边 | 调整侧栏宽度（38 ~ 80 mm） |
| 单击 ✕ | 收进托盘 |

> 标题栏只有 ✕ 一个按钮 —— 之前有三个圆点（红/黄/绿），但侧栏只会在放映时出现，
> 绿点"正在放映"指示完全冗余；黄点的"换边"功能拖动+托盘菜单都能做；
> 三个点加 ✕ 正好把标题栏铺满，手指根本拖不动。

### 关闭 / 隐藏侧栏

注意区分两个概念：

- **✕ / 托盘「显示·隐藏」= 收进托盘**（程序还在跑，下次放映照样自动弹出）
- **托盘右键 → 退出侧栏 = 真正结束程序**

出口一共这几条：

1. **点右上角 ✕** → 收进托盘
2. **托盘图标**（右下角）**双击** → 显示 / 隐藏切换
3. **托盘图标右键 → 退出侧栏（关闭本程序）** → 彻底退出
4. **托盘图标右键 → 当前状态 / PID** → 显示侧栏 PID、贴边位置、是否在放映
5. 双击 `关闭侧栏.cmd`，或命令行 `PowerPointSidebar.exe --kill` → 找不到托盘图标时的兜底

> 侧栏有意**不出现在任务栏**（否则全屏放映时会把任务栏带出来），所以出口在托盘和右下角 ✕，不在任务栏。

---

## 三、六个按钮都是什么

| 按钮 | 调用的 PowerPoint 接口 | 说明 |
|---|---|---|
| 上一页 | `SlideShowView.Previous()` | 上一页 / 上一个动画 |
| 下一页 | `SlideShowView.Next()` | 下一页 / 下一个动画 |
| 笔 | `PointerType = 2` + `PointerColor.RGB` | 用「笔色」画，颜色可自定义（点小色块换色） |
| 橡皮 | `PointerType = 5` | 一笔一笔擦（**双击 = 清空本页笔迹**） |
| 结束放映 | `SlideShowView.Exit()` | 结束放映回到普通视图（会先弹确认框） |

> ⚠️ 「结束放映」是**结束幻灯片放映**，不是关闭侧栏。关闭侧栏请点右上角 ✕（收托盘）。

### 之前有个「荧光笔」按钮，已删除

理由：

1. **它挂着错误的承诺** —— PowerPoint 的 COM 接口只暴露一种笔，没有真正的半透明高亮。
   「荧光笔」实际是黄色普通笔，叫这个名字会让人期待"半透明粗高亮"，然后失望。
2. **笔色已经能改** —— 点「笔」右侧的色块就能选任意亮黄色，要"荧光效果"换色就行。
3. **侧栏更紧凑** —— 删掉一个按钮省掉一整行（11 mm），侧栏从 6 行变 5 行。

`PptController.UseHighlighter()` 这个 COM 方法保留着，将来若要做 in-process 加载项（半透明画线）还用得上。

### ⚠️ 橡皮有个 PowerPoint 自身的限制（重要）

实测结论（本机 Office 16 x64）：

| 场景 | `PointerType = 5` 的结果 |
|---|---|
| 页面上**已有笔迹** | 读回 5 ✅ 橡皮生效 |
| 页面**还是空白** | 读回 4 ❌ PowerPoint **静默忽略** |

**PowerPoint 在页面上没有任何墨迹时会直接忽略橡皮指令**（0/1/2/3/4 则任何时候都能写入）。

侧栏已经处理了这一点：切橡皮后会回读确认，如果没切上会明确提示
「这一页还没有笔迹，先用『笔』画两笔再点橡皮」，而不是让你感觉"点了没反应"。

### 🎨 笔色（可自定义）

**「笔」的颜色可以改**：

1. 点「笔」按钮**右侧的小色块**（不是点按钮本身）
2. 弹出调色板 → 点一个颜色 → 立即生效并记住
3. 调色板里还有「恢复默认」

要点：
- 点按钮**本体**是切换工具；点**小色块**才是换颜色，两者不会互相干扰
- 选完颜色如果正在放映，会**立刻**套用到 PowerPoint 的笔上
- 颜色保存在 `%APPDATA%\PowerPointSidebar\settings.txt`，重启侧栏后仍然生效
  （不放在 `bin\` 目录，因为重新构建会清掉）
- 也可以从**托盘右键 → 「笔的颜色…」**打开调色板

**为什么只有 15 个预设色、不能选任意 RGB**：侧栏带 `WS_EX_NOACTIVATE`（为了不抢放映时的键盘焦点），
这导致覆盖层里的输入框**拿不到键盘焦点**，hex 输入这类交互直接不可用；
而系统取色器是需要激活的模态对话框，在全屏放映上方弹出也不可靠。
15 个预设色对"放映时标注"够用了。真要任意色，直接编辑上面那个 `settings.txt` 即可。

### ⚠️ 「笔」做不到的事（重要，别误会）

**PowerPoint 的 COM 接口只暴露一种笔**，没有真正的半透明高亮，也**无法通过 COM 调整笔头粗细**。

所以这里的「笔」= 普通粗细的实色笔，不是真正那种半透明的宽高亮带。

想要真正的半透明荧光笔，只有两条路：

1. 不用 COM 的笔，改成用 `Shapes.AddLine` 在幻灯片上画半透明形状 —— 但需要实时跟踪鼠标轨迹，独立进程做不到，必须做成 in-process 加载项
2. 手动在 PowerPoint 里把笔头调粗（右键放映画面 → 指针选项 → 笔或荧光笔）

这个小限制在下面的「已知限制」里也列了。

### 关于焦点（已彻底解决）

侧栏窗口带 `WS_EX_NOACTIVATE` 扩展样式，**点击它完全不夺取键盘焦点**——焦点始终留在 PowerPoint，
所以放映时用翻页笔、方向键翻页不会受任何影响。

（第一版没有这个样式，靠"每次操作后把焦点交回 PowerPoint"来补救，会有一瞬间的焦点抖动；
现在从根上不抢焦点了，`SetForegroundWindow` 只作为兜底保留。）

### 为什么不占任务栏

侧栏带 `WS_EX_TOOLWINDOW` 样式且 `ShowInTaskbar = false`，**不会出现在任务栏，也不出现在 Alt+Tab**。

这一点很重要：全屏放映时 Windows 会隐藏任务栏，但如果侧栏是个"普通任务栏窗口"，
Windows 就会**把任务栏重新显示出来**，盖在幻灯片上。第一版就是踩了这个坑
（当时为了给你一个"关得掉"的出口把 `ShowInTaskbar` 设成了 `true`），
现在出口改到**托盘图标**，既关得掉又不干扰放映。

---

## 四、已知限制

1. **没有真正的半透明笔**（PowerPoint COM 限制，只能用实色普通笔）
2. **不会自动启动 PowerPoint**，必须先自己打开
3. 只在**放映状态**（有 SlideShowWindow）下按钮才生效；普通编辑视图下点击会提示「当前没有幻灯片放映窗口」
4. 关闭 PowerPoint 后侧栏会自动解绑，重新打开 PowerPoint 并放映后会**自动重新挂载**（每 0.6 秒轮询一次）
5. 需要本机已安装 PowerPoint 桌面版（本机为 Office 16 / x64，已确认可用）
6. 侧栏刻意**不在任务栏留图标**，所以"找不到它"时请去**右下角系统托盘**找
   （或者直接 `PowerPointSidebar.exe --kill`）

---

## 五、代码结构

| 文件 | 作用 |
|---|---|
| `Program.cs` | 入口。命令行参数（`--show` / `--selftest` / `--kill`）、高 DPI、全局异常兜底 |
| `MainForm.cs` | 侧栏窗体：Win32 扩展样式、贴边、托盘、内嵌提示条与调色板、`SidebarButton` 自绘按钮 |
| `PptController.cs` | COM 层。**后期绑定**（`Type.GetTypeFromProgID` + `InvokeMember`），不依赖 Office PIA |
| `SettingsStore.cs` | 极简设置存储：读写 `%APPDATA%\PowerPointSidebar\settings.txt` 里的笔色 |
| `IconFactory.cs` | 用 GDI+ 现画图标，不引外部图片 |

### 关键实现点

**窗口样式（三个都要，缺一不可）**

```csharp
protected override CreateParams CreateParams {
    get {
        var cp = base.CreateParams;
        cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE  点击不抢焦点
        cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW  不进任务栏/Alt+Tab
        return cp;
    }
}
```

**维持置顶**：PowerPoint 的全屏放映窗口自己也是置顶的，会把侧栏压下去。
所以有个 700ms 的定时器周期性重新断言：

```csharp
SetWindowPos(Handle, HWND_TOPMOST, 0,0,0,0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
```

**自动弹出 / 收纳**：一个 600ms 的定时器轮询 `IsSlideshowRunning`，
状态翻转时 `Show()`（并 `ApplyDock()` 贴边）或 `Hide()`（收托盘）。

**布局**：按钮把标题栏以下的剩余高度**平均分配**，所以贴满屏幕高度时按钮会被拉高铺满，
不会底部留一大块空白。

### 几个容易踩的坑（都已在本项目修掉，改代码时注意）

1. ★★★ **Office 集合的 `Item` 不能一律当属性调用。**
   `SlideShowWindows.Item(1)` 用 `GetProperty` 调用会直接抛
   `DISP_E_MEMBERNOTFOUND (0x80020003)`，**必须用 `InvokeMethod`**。
   第一版踩了这个坑，导致**六个按钮全都报"找不到成员"**。
   现在 `GetCollectionItem()` 先按方法试、失败再按属性试。

2. ★★ **同一个坑会在别处重演。** 写诊断工具时我又写了一次 `Get(windows, "Item", 1)`，
   结果导出放映失败——别以为写过一次就不会再犯。

3. ★★★ **`PpSlideShowPointerType` 的值必须背准，尤其橡皮是 `5`。**
   | 值 | 含义 |
   |---|---|
   | 0 | None |
   | 1 | Arrow |
   | 2 | **Pen** |
   | 3 | **AlwaysHidden**（指针永久隐藏）← 别当成橡皮 |
   | 4 | AutoArrow |
   | **5** | **Eraser** |

   第二版把橡皮写成 `3`，于是点「橡皮」实际是**把指针藏起来**，
   用户看到的症状是"点击没反应，双击才行"（双击走的是 `EraseDrawing`，有可见效果）。

4. ★★ **橡皮有前置条件：页面上必须已经有墨迹。**
   实测：空白页写 `PointerType = 5` 会被 **静默忽略**（读回仍是 4 AutoArrow）；
   页面上有墨迹时才能写入成功（读回 5）。0/1/2/3/4 则任何时候都能写。
   所以 `UseEraser()` 改成**写入后回读确认**并返回 bool，UI 据此给出解释，
   否则用户只会看到"点了没反应"。

5. ★★ **`SlideShowView.PointerColor` 是只读的 `ColorFormat` 对象，不是整数。**
   正确写法是 `PointerColor.RGB = <BGR 整数>`，而不是 `PointerColor = <整数>`。
   虽说后者不报错，但实测**真的不生效**。现在 `SetPointerColor()` 先走 `.RGB`，
   失败再退回直接赋值做兜底。
   （旁证：早期诊断里读 `PointerColor` 返回的是 `System.__ComObject`，正是 ColorFormat。）

6. **`WS_EX_NOACTIVATE` 的副作用不止一个**：
   ① 模态 `MessageBox` 无法正常激活（按钮点不动）→ 本项目所有提示改用**内嵌覆盖层**；
   ② **文本框拿不到键盘焦点** → 所以调色板只能做点选式色块，做不了 hex 输入。

7. ★★ **无边框 + 无任务栏窗口必须给用户留可见出口。** 第一版 `ShowInTaskbar=false`
   且关闭点只有 10px 圆点，用户直接"关不掉了"；第二版改成 `true`，又导致全屏放映时
   **把系统任务栏带出来**。正解是**托盘图标**（`NotifyIcon`）——既不干扰全屏，又永远找得到。

8. **不要手写 `Properties/AssemblyInfo.cs`**：SDK 风格项目会自动生成程序集特性，手写会撞 `CS0579 特性重复`

9. **高 DPI 只在 `Program.Main` 里设**：`app.manifest` 里再写 `<dpiAware>` 会撞 `WFO0003` 警告

10. **自定义控件的 public 属性**会被 WinForms 分析器报 `WFO1000`（要求配置 designer 序列化）。
    文件顶部用 `#pragma warning disable WFO1000` 统一关掉

11. **`Timer` 有歧义**：`using System.Windows.Forms;` 时会和 `System.Threading.Timer` 冲突
    （`CS0104`），要写全 `System.Windows.Forms.Timer`

12. **`Application.HWND` 在没有打开演示文稿时返回 0**，这时要退回用 Win32 `EnumWindows`
    枚举 `POWERPNT` 进程的顶层窗口，取面积最大且无 owner 的那个

13. **不要用 `Application.Run(form)`**：它会强制 `Show()` 一次，启动瞬间闪一个窗口。
    侧栏平时应该安静待在托盘，所以用 `Application.Run(new ApplicationContext(form))`
    再手动 `form.Start()`

14. **重建前先停掉正在运行的实例**，否则 exe 被锁会报 `MSB3027 / MSB3021`。
    另外如果进程是用后台 shell 拉起来的，要连那个 shell 任务一起停，
    光 `taskkill` 可能释放不掉（本项目就遇到过 `HasExited=True` 但文件仍被锁的情况）。

---

## 六、怎么自己复现/验证 COM 问题（开发笔记）

侧栏按钮报错时，**不要靠猜**。实测过的定位方法：

因为本机没有 Office PIA，PowerShell 的 `New-Object -ComObject PowerPoint.Application`
也不方便测「和 PptController 完全一致的反射调用方式」，所以最可靠的做法是
**建一个临时控制台工程，把 `PptController.cs` 直接 `<Compile Include>` 进去**，
然后自建一个临时文稿跑放映、逐个调用生产方法：

```xml
<ItemGroup>
  <Compile Include="...\PowerPointSidebar\PptController.cs" />
</ItemGroup>
```

这样测的就是**真实代码路径**，而不是复刻一份逻辑（复刻的往往和真实的不一致，
这次第一版诊断就是因为复刻逻辑才没第一时间定位到）。

实测结论（2026-09-25，Office 16 x64）：

```
A) Item  (GetProperty)   : FAIL  0x80020003 DISP_E_MEMBERNOTFOUND
B) Item  (InvokeMethod)  : OK
get View (GetProperty)   : OK
读 PointerType / PointerColor (GetProperty) : OK
写 PointerType = 2 / PointerColor (SetProperty) : OK
Next() / Previous() / EraseDrawing() / Exit() (InvokeMethod) : OK
```

### 窗口样式与贴边实测

写一个只读的 `EnumWindows` 工具，读侧栏窗口的 `GWL_EXSTYLE` 和矩形，实测结果：

```
窗口标题 : PPT 侧栏
矩形     : (1160,0) - (1280,752)  尺寸 120x752
ExStyle  : 0x08010088
  WS_EX_NOACTIVATE (不抢焦点)  : True
  WS_EX_TOOLWINDOW (不进任务栏) : True
  WS_EX_TOPMOST    (置顶)      : True
  WS_EX_APPWINDOW  (进任务栏)  : False
贴边判断 : 屏幕[\\.\DISPLAY1] 贴右边缘 整屏高度
```

### 自动弹出 / 收纳实测（三阶段）

用一个临时工具起一个 14 秒的真放映，中途采样三次：

| 阶段 | 窗口可见 | 判定 |
|---|---|---|
| A 无放映 | `False` | 已缩进托盘 ✓ |
| B 放映进行中 | `True`（贴右边缘·整屏高） | 自动弹出并贴边 ✓ |
| C 放映结束后 | `False` | 自动收回托盘 ✓ |

三个阶段里 `WS_EX_NOACTIVATE / TOOLWINDOW / TOPMOST` 与 `APPWINDOW=False` 都保持正确，
说明样式在显示/隐藏过程中没有丢失。

临时文稿在测试结束时 `Close()` 不保存，实测 `清理后剩余文稿数` 与测试前一致，
**用户原有文稿未被改动**。

---

## 七、重新构建
