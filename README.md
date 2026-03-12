# HOW-TO 灏图品测 · 智能报价系统

## 第一次部署（只需做一次）

### 第一步：安装 Node.js
打开 https://nodejs.org → 点击左边绿色大按钮下载 → 安装（一路下一步）

安装完成后打开"终端"（Mac）或"命令提示符"（Windows），输入：
```
node -v
```
如果显示版本号（如 v20.xx.x）说明安装成功。

---

### 第二步：注册 GitHub 账号
打开 https://github.com → Sign up → 注册一个账号（记住用户名和密码）

---

### 第三步：注册 Vercel 账号
打开 https://vercel.com → Sign Up → 选择"Continue with GitHub" → 授权

---

### 第四步：把项目文件夹上传到 GitHub

1. 打开 https://github.com/new
2. Repository name 填：`hut-quoter`
3. 点绿色"Create repository"按钮
4. 在终端里，进入本项目文件夹，依次运行：

```bash
cd hut-quoter          # 进入文件夹（根据实际路径调整）
npm install            # 安装依赖（只需一次，会出现进度条）
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/你的GitHub用户名/hut-quoter.git
git push -u origin main
```

> ⚠️ 把"你的GitHub用户名"替换成你自己的用户名

---

### 第五步：在 Vercel 部署

1. 打开 https://vercel.com/dashboard
2. 点击"Add New → Project"
3. 找到"hut-quoter"，点"Import"
4. 所有设置保持默认，直接点"Deploy"
5. 等待约1分钟，出现"Congratulations!"即成功
6. 点击显示的链接（如 https://hut-quoter-xxx.vercel.app）即可访问

把这个链接发给同事，任何人都可以用。

---

## 以后更新代码

每次修改了 App.jsx 之后，在终端运行：

```bash
git add .
git commit -m "更新内容描述"
git push
```

Vercel 会自动检测到代码变化，1分钟内自动更新线上版本。

---

## 文件结构说明

```
hut-quoter/
├── index.html          # 入口 HTML（不需要改）
├── package.json        # 项目配置（不需要改）
├── vite.config.js      # 构建配置（不需要改）
└── src/
    ├── main.jsx        # 启动文件（不需要改）
    ├── index.css       # 全局样式（不需要改）
    └── App.jsx         # ← 主要逻辑，改功能改这个文件
```
