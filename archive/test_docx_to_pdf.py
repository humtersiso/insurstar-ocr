import os
import glob
import multiprocessing
import time

def wrapper(q, func, args):
    try:
        result = func(*args)
        q.put(result)
    except Exception as e:
        q.put(e)

def run_with_timeout(func, args=(), timeout=60):
    q = multiprocessing.Queue()
    p = multiprocessing.Process(target=wrapper, args=(q, func, args))
    p.start()
    p.join(timeout)
    if p.is_alive():
        p.terminate()
        p.join()
        print(f"[timeout] {func.__name__} 超過 {timeout} 秒，強制中止")
        return False
    if not q.empty():
        result = q.get()
        if isinstance(result, Exception):
            print(f"[{func.__name__}] 發生例外: {result}")
            return False
        return result
    return False

def try_docx2pdf(docx_path, pdf_path):
    try:
        from docx2pdf import convert
        import pythoncom
        pythoncom.CoInitialize()
        convert(docx_path, pdf_path)
        pythoncom.CoUninitialize()
        return os.path.exists(pdf_path)
    except Exception as e:
        print(f"docx2pdf 失敗: {e}")
        return False

def try_win32com(docx_path, pdf_path):
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close()
        word.Quit()
        pythoncom.CoUninitialize()
        return os.path.exists(pdf_path)
    except Exception as e:
        print(f"win32com 失敗: {e}")
        return False

def try_pypandoc(docx_path, pdf_path):
    try:
        import pypandoc
        output = pypandoc.convert_file(docx_path, 'pdf', outputfile=pdf_path)
        return os.path.exists(pdf_path)
    except Exception as e:
        print(f"pypandoc 失敗: {e}")
        return False

def convert_docx_to_pdf(docx_path):
    methods = [
        (try_docx2pdf, 'docx2pdf'),
        (try_win32com, 'win32com'),
        (try_pypandoc, 'pypandoc'),
    ]
    results = []
    print(f"\n==== 測試 {os.path.basename(docx_path)} ====")
    for func, method_name in methods:
        pdf_path = docx_path.replace('.docx', f'_{method_name}.pdf')
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        start = time.time()
        success = run_with_timeout(func, (docx_path, pdf_path), timeout=60)
        elapsed = time.time() - start
        results.append({
            'method': method_name,
            'success': success,
            'time': elapsed,
            'pdf_path': pdf_path if success else None
        })
        print(f"{method_name}: {'成功' if success else '失敗'}，耗時 {elapsed:.2f} 秒")
    print("---- 結果彙整 ----")
    for r in results:
        print(f"方法: {r['method']}, 結果: {'成功' if r['success'] else '失敗'}, 耗時: {r['time']:.2f} 秒, PDF: {r['pdf_path'] if r['pdf_path'] else '-'}")

if __name__ == "__main__":
    docx_files = glob.glob('property_reports/*.docx')
    if not docx_files:
        print("找不到任何 .docx 檔案")
    for docx_path in docx_files:
        convert_docx_to_pdf(docx_path) 