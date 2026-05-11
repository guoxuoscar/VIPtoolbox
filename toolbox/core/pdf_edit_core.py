# -*- coding: utf-8 -*-
"""
PDF 权限清理核心：自动检测加密/禁止修改，重写为无限制可编辑副本。
从 5.4pdf编辑 项目提取，适配工具箱环境。
"""
from __future__ import annotations

import io
import os
from typing import Any

import fitz

_SAVE_ENCRYPT_PERMISSIONS = (
    fitz.PDF_PERM_ACCESSIBILITY
    | fitz.PDF_PERM_PRINT
    | fitz.PDF_PERM_PRINT_HQ
    | fitz.PDF_PERM_MODIFY
    | fitz.PDF_PERM_COPY
    | fitz.PDF_PERM_ANNOTATE
    | fitz.PDF_PERM_FORM
    | fitz.PDF_PERM_ASSEMBLE
)


def _needs_security_strip(doc: fitz.Document) -> bool:
    if doc.is_encrypted:
        return True
    return (doc.permissions & fitz.PDF_PERM_MODIFY) == 0


def open_pdf_editable_copy(data: bytes) -> tuple[fitz.Document, dict[str, Any]]:
    """
    打开 PDF 并尽量转为「无打开密码、可自由编辑」的内存副本。
    仅支持「用系统默认方式能直接打开、不要求输入密码」的文件。
    若检测到加密或禁止修改等限制，则重写为不加密的新文档再返回。
    """
    doc = fitz.open(stream=data, filetype="pdf")
    if doc.needs_pass:
        if not doc.authenticate(""):
            doc.close()
            raise ValueError(
                "此 PDF 设置了打开密码（打开时必须手动输入密码）。"
                "本工具只处理「双击即可打开、不要求输入密码」的 PDF；"
                "请先用其它软件去掉打开密码。"
            )
    if not _needs_security_strip(doc):
        return doc, {"security_stripped": False}
    buf = io.BytesIO()
    doc.save(buf, garbage=3, deflate=True)
    doc.close()
    data2 = buf.getvalue()
    doc2 = fitz.open(stream=data2, filetype="pdf")
    return doc2, {"security_stripped": True}


def open_pdf_editable_path(path: str) -> tuple[fitz.Document, dict[str, Any]]:
    """从磁盘路径打开并转为可编辑副本。"""
    with open(path, "rb") as f:
        return open_pdf_editable_copy(f.read())


def save_document(doc: fitz.Document, out_path: str) -> bytes:
    """
    保存文档为无加密 PDF 到磁盘。
    返回写入的字节数（便于日志记录）。
    """
    buf = io.BytesIO()
    doc.save(buf, deflate=True, garbage=3)
    data = buf.getvalue()
    with open(out_path, "wb") as f:
        f.write(data)
    return data
