#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
GitHub 仓库一键初始化工具
用法：
  python scripts/init_github.py                  # 交互式输入仓库名
  python scripts/init_github.py --name 仓库名     # 直接指定仓库名

首次使用需要登录 GitHub（一次性的，之后不用再登）：
  运行后按提示打开浏览器，输入设备码即可。
"""
import os
import sys
import subprocess
import argparse

# 确保 gh 在 PATH 中
GH_PATH = r"C:\Program Files\GitHub CLI"
if GH_PATH not in os.environ.get("PATH", ""):
    os.environ["PATH"] = GH_PATH + os.pathsep + os.environ["PATH"]

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def run(cmd, capture=True):
    """运行命令，返回 (成功?, 输出)"""
    try:
        r = subprocess.run(
            cmd, capture_output=capture, text=True, encoding="utf-8",
            cwd=PROJECT_DIR, timeout=60,
        )
        ok = r.returncode == 0
        return ok, r.stdout.strip() if capture else r.stdout
    except Exception as e:
        return False, str(e)


def check_gh():
    """检查 gh 是否可用"""
    ok, out = run(["gh", "--version"])
    if not ok:
        print("[X] GitHub CLI (gh) 未安装，正在自动安装...")
        ok2, _ = run(["winget", "install", "--id", "GitHub.cli"], capture=False)
        if not ok2:
            print("[X] 自动安装失败，请手动安装：")
            print("    https://cli.github.com/")
            return False
        ok, out = run(["gh", "--version"])
    print(f"[OK] GitHub CLI: {out.split(chr(10))[0]}")
    return True


def check_auth():
    """检查是否已登录"""
    ok, out = run(["gh", "auth", "status"])
    if ok:
        print("[OK] 已登录 GitHub")
        return True
    # 未登录，引导登录
    print("[!] 未登录 GitHub，需要你完成一次登录（仅首次）")
    print()
    try:
        r = subprocess.run(
            ["gh", "auth", "login", "-w"],
            cwd=PROJECT_DIR,
            timeout=120,
        )
        return r.returncode == 0
    except subprocess.TimeoutExpired:
        print("[X] 登录超时")
        return False
    except Exception as e:
        print(f"[X] 登录失败: {e}")
        return False


def create_repo(name, description="", private=False, push_existing=False):
    """创建 GitHub 仓库"""
    visibility = "private" if private else "public"

    # 创建仓库
    cmd = ["gh", "repo", "create", name, "--" + visibility, "--source=.", "--push"]
    if description:
        cmd += ["--description", description]

    print(f"[.] 正在创建仓库 {name} ({visibility})...")
    ok, out = run(cmd, capture=False)
    if ok:
        print(f"[OK] 仓库创建成功！")
        print(f"     https://github.com/{name}")
        return True
    else:
        print(f"[X] 创建失败")
        return False


def main():
    # 先 cd 到项目目录
    os.chdir(PROJECT_DIR)

    parser = argparse.ArgumentParser(description="GitHub 仓库一键初始化")
    parser.add_argument("--name", help="仓库名（例如 guoxuoscar/MyProject）")
    parser.add_argument("--desc", help="仓库描述", default="")
    parser.add_argument("--private", action="store_true", help="创建私有仓库")
    parser.add_argument("--push-existing", action="store_true",
                        help="推送到已有仓库（跳过创建）")
    args = parser.parse_args()

    # 1) 检查 gh
    if not check_gh():
        sys.exit(1)

    # 2) 检查登录
    if not check_auth():
        sys.exit(1)

    # 3) 获取或创建仓库
    if args.push_existing:
        # 已有远程仓库，直接推送
        ok, out = run(["git", "remote", "-v"])
        if ok and out:
            print(f"[OK] 远程仓库已存在：\n{out}")
        else:
            print("[!] 没有远程仓库，请先指定 --name 创建新仓库")
            sys.exit(1)
    else:
        name = args.name
        if not name:
            # 从 git config 获取用户名
            ok, user = run(["git", "config", "user.name"])
            default_name = user.strip() if ok and user else "用户名"
            print("== GitHub 仓库创建 ==")
            print(f"GitHub 用户: {default_name}")
            name_input = input("仓库名 (例如 MyProject): ").strip()
            if not name_input:
                print("[X] 仓库名不能为空")
                sys.exit(1)
            if "/" in name_input:
                name = name_input  # 完整路径
            else:
                name = f"{default_name}/{name_input}"

        private = args.private
        if not args.private:
            priv = input("私有仓库? (y/N): ").strip().lower()
            private = priv == "y"

        desc = args.desc
        if not desc:
            desc = input("描述 (可选): ").strip()

        create_repo(name, desc, private)


if __name__ == "__main__":
    main()
