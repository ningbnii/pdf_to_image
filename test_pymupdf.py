"""
测试PyMuPDF打开PDF文件的各种方式
"""
import sys
import fitz  # PyMuPDF


def test_pymupdf():
    print(f"PyMuPDF 版本: {fitz.__version__}")
    print(f"Python 版本: {sys.version}")

    print("\n检查fitz模块的属性:")
    methods = [attr for attr in dir(fitz) if not attr.startswith('__')]
    print(f"模块方法数量: {len(methods)}")
    print(f"部分方法: {methods[:10]}")

    print("\n检查是否有open方法:")
    if hasattr(fitz, 'open'):
        print("fitz.open 存在")
    else:
        print("fitz.open 不存在")

    print("\n检查是否有Document类:")
    if hasattr(fitz, 'Document'):
        print("fitz.Document 存在")
    else:
        print("fitz.Document 不存在")

    # 尝试不同方式打开PDF文件
    # 注意：需要有一个示例PDF文件用于测试
    pdf_path = "example.pdf"
    import os
    if not os.path.exists(pdf_path):
        print(f"\n{pdf_path} 不存在，无法进行打开PDF的测试")
        return

    print("\n尝试用不同方式打开PDF:")

    try:
        if hasattr(fitz, 'open'):
            print("尝试 fitz.open()...")
            doc = fitz.open(pdf_path)
            print(f"成功! 页数: {len(doc)}")
            doc.close()
    except Exception as e:
        print(f"fitz.open() 失败: {e}")

    try:
        if hasattr(fitz, 'Document'):
            print("尝试 fitz.Document()...")
            doc = fitz.Document(pdf_path)
            print(f"成功! 页数: {len(doc)}")
            doc.close()
    except Exception as e:
        print(f"fitz.Document() 失败: {e}")


if __name__ == "__main__":
    test_pymupdf()
