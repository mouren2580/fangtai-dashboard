# 方太西北服务产品看板 · 家里电脑一键同步说明

本说明教你如何在**家里（或任何一台新）电脑**上，用最新月报 Excel 自己刷新看板并发布上线。

> 看板在线地址不变：`https://mouren2580.github.io/fangtai-dashboard/`
> 同步脚本 `sync.py` 已做成**自包含**：不依赖办公室的盘符路径、也不依赖 WorkBuddy 技能，复制整个文件夹即可用。

---

## 一、把项目文件夹整体拷到家里电脑

从办公室电脑（或让同事/网盘）拿到这个文件夹，**原样整体复制**到家里电脑任意位置（如 `D:\fangtai-dashboard\`）。文件夹里至少要包含：

```
dashboard.html          ← 看板本体（sync.py 只刷新里面的数据）
sync.py                 ← 一键同步脚本（自包含）
build_dashboard.py      ← 核心聚合逻辑（已随包提供，无需装技能）
requirements.txt        ← Python 依赖清单
sync.bat               ← Windows 一键脚本
.gh_token              ← 你自己建的令牌文件（见第三步，勿从办公室复制我的）
```

> ⚠️ `.gh_token` **不要**从办公室电脑直接复制过来——那是办公室的密钥。家里电脑要**自己新建**一个（第三步）。

---

## 二、安装 Python 和依赖

1. 到 https://www.python.org/downloads/ 下载安装 **Python 3.11 或更新版本**。
   - **关键**：安装时务必勾选 **`Add python.exe to PATH`**（添加到环境变量）。
2. 安装完成后，打开命令行（Win+R → 输入 `cmd` → 回车），在**项目文件夹内**执行：
   ```bat
   pip install -r requirements.txt
   ```
   看到 `Successfully installed openpyxl-...` 即成功。

---

## 三、准备 GitHub 发布令牌（只需做一次）

看板发布到 GitHub Pages 需要一把"钥匙"（个人访问令牌）。你自己生成一把：

1. 登录 GitHub → 右上角头像 → **Settings** → 左侧 **Developer settings** → **Personal access tokens** → **Tokens (classic)**。
2. 点 **Generate new token (classic)**。
3. Note 随便写（如 `fangtai-home`）；Expiration 选长一点（如 90 天或 No expiration）。
4. **勾选 `repo`（整组）** → 拉到底点 **Generate token**。
5. 复制生成的令牌，形如 `ghp_xxxxxxxxxxxxxxxx`。**只显示这一次，妥善保存**。
6. 在项目文件夹里**新建一个文件叫 `.gh_token`**（注意前面有个点），用记事本打开，把令牌**单独一行粘贴进去**保存。**不要带空格/换行**。

> 之后每次同步都会自动读取这个文件，不用重复生成。

---

## 四、把最新月报 Excel 放好

把最新的 `2026年8月西北服务产品.xlsx`（或对应月份）放到项目文件夹内，记好文件名。
Excel 里必须包含这四个表（和办公室版一致）：
`服务产品收入统计`、`CSM服务项目`、`CSM配件`、`WMS网点买断配件明细`。

---

## 五、一键同步

**最简单**：双击 `sync.bat`，按提示输入 Excel 文件名和截止日即可。

**或命令行**（在项目文件夹内）：
```bat
python sync.py --excel "2026年8月西北服务产品.xlsx" --date 2026-08-16
```

参数说明：
- `--excel`：最新月报 Excel 路径（放文件夹内就只写文件名）。
- `--date`：**数据截止日 = 你拿到 Excel 那天的前一天**，格式 `YYYY-MM-DD`。
  - 例如 8/17 导出的数据 → 截止日写 `2026-08-16`。这是铁律，别写错。
- 令牌默认读 `.gh_token`；也可临时用 `--token ghp_xxx`。
- 想**先只看效果、不发布**，加 `--no-deploy`：
  ```bat
  python sync.py --excel "2026年8月西北服务产品.xlsx" --date 2026-08-16 --no-deploy
  ```
- 只有当确实拿到新数据、要对外更新页面时，才加 `--bump-stamp`（会推进页面上的"数据更新时间"）。默认不加，符合"未确认更新不滚动链接日期"的约定。

运行成功会打印：网点数、总实际金额、工程师数、延保合计、周数据等，最后显示 `OK 发布成功`。

---

## 六、查看

浏览器打开 `https6://mouren2580.github.io/fangtai-dashboard/`，按 **Ctrl+Shift+R** 硬刷新即可看到新数据。

---

## 七、常见问题

| 现象 | 排查 |
|------|------|
| `python` 不是内部命令 | 第二步没勾 PATH；重装 Python 并勾选 Add to PATH，或重启命令行。 |
| `No module named openpyxl` | 没跑 `pip install -r requirements.txt`，或装到了别的 Python。 |
| `未提供令牌` / 发布失败 401 | `.gh_token` 没建好或内容有空格换行；令牌需有 `repo` 权限。 |
| `找不到 Excel` | `--excel` 路径写错，或文件名不对。 |
| 延保/周数据为空 | Excel 里 `CSM服务项目`/`CSM配件`/`WMS网点买断配件明细` 表名不一致，或该月确实无数据。 |
| 家里打不开 github.io | 国内家庭宽带访问 GitHub Pages 不稳定，可换无痕窗口/手机热点，或让我把看板迁到国内托管。 |

---

## 八、备注

- `sync.py` 刷新的是**当前基准月**（8 月）的主数据 + 服务工程师 + 延保 + 周数据；1–7 月为历史冻结数据，不在每次同步范围内。
- 切换新月份（如 9 月）需注入新月份数据，步骤较复杂，建议仍交回办公室让我处理，或单独再说。
- 本文件夹与办公室是**同一套逻辑**，两边都能跑；谁最后同步，线上就是谁的数据。
