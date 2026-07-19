import requests
import time
import os
import io
from datetime import datetime
import openpyxl

BASE_URL = "http://localhost:8000"
TEST_FILE_PATH = os.path.join("data", "protected_file.xlsx")

def log_test_result(test_name, method, endpoint, status_code, response_time, response_data, passed):
    print(f"\n{'='*70}")
    print(f"测试用例: {test_name}")
    print(f"请求方法: {method}")
    print(f"请求端点: {endpoint}")
    print(f"状态码: {status_code}")
    print(f"响应时间: {response_time:.3f}秒")
    print(f"测试结果: {'✅ 通过' if passed else '❌ 失败'}")
    if isinstance(response_data, dict) or isinstance(response_data, list):
        print(f"响应数据: {response_data}")
    elif isinstance(response_data, bytes):
        print(f"响应数据: 二进制数据 ({len(response_data)} bytes)")
    elif response_data:
        print(f"响应数据: {response_data}")
    print(f"{'='*70}")
    return {
        "test_name": test_name,
        "method": method,
        "endpoint": endpoint,
        "status_code": status_code,
        "response_time": response_time,
        "response_data": response_data if isinstance(response_data, (dict, list, str)) else f"binary data ({len(response_data)} bytes)",
        "passed": passed
    }

def test_health_check():
    start_time = time.time()
    try:
        response = requests.get(f"{BASE_URL}/health")
        response_time = time.time() - start_time
        passed = response.status_code == 200 and response.json().get("status") == "ok"
        return log_test_result("健康检查 - GET /health", "GET", "/health", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("健康检查 - GET /health", "GET", "/health", -1, response_time, str(e), False)

def test_unprotect_valid_file():
    start_time = time.time()
    try:
        with open(TEST_FILE_PATH, "rb") as f:
            files = {"file": ("protected_file.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            response = requests.post(f"{BASE_URL}/unprotect", files=files)
        response_time = time.time() - start_time
        
        checks = []
        checks.append(response.status_code == 200)
        
        content_disposition = response.headers.get("Content-Disposition", "")
        checks.append("unprotected_protected_file.xlsx" in content_disposition)
        checks.append("attachment" in content_disposition)
        
        content_type = response.headers.get("Content-Type", "")
        expected_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        checks.append(expected_type in content_type)
        
        try:
            output_stream = io.BytesIO(response.content)
            wb = openpyxl.load_workbook(output_stream)
            for sheet in wb.worksheets:
                checks.append(sheet.protection.sheet == False)
            print(f"  ✅ 验证通过: 返回文件包含 {len(wb.worksheets)} 个工作表，所有工作表保护已禁用")
        except Exception as e:
            print(f"  ❌ 文件验证失败: {str(e)}")
            checks.append(False)
        
        passed = all(checks)
        return log_test_result("POST /unprotect - 正常上传受保护的Excel文件", "POST", "/unprotect", response.status_code, response_time, response.content, passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("POST /unprotect - 正常上传受保护的Excel文件", "POST", "/unprotect", -1, response_time, str(e), False)

def test_unprotect_empty_filename():
    start_time = time.time()
    try:
        files = {"file": ("", b"", "application/octet-stream")}
        response = requests.post(f"{BASE_URL}/unprotect", files=files)
        response_time = time.time() - start_time
        passed = response.status_code == 400
        return log_test_result("POST /unprotect - 上传空文件名", "POST", "/unprotect", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("POST /unprotect - 上传空文件名", "POST", "/unprotect", -1, response_time, str(e), False)

def test_unprotect_txt_file():
    start_time = time.time()
    try:
        files = {"file": ("test.txt", b"This is a test text file.", "text/plain")}
        response = requests.post(f"{BASE_URL}/unprotect", files=files)
        response_time = time.time() - start_time
        passed = response.status_code == 400
        return log_test_result("POST /unprotect - 上传非xlsx文件(.txt)", "POST", "/unprotect", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("POST /unprotect - 上传非xlsx文件(.txt)", "POST", "/unprotect", -1, response_time, str(e), False)

def test_unprotect_xls_file():
    start_time = time.time()
    try:
        files = {"file": ("test.xls", b"This is a test xls file content.", "application/vnd.ms-excel")}
        response = requests.post(f"{BASE_URL}/unprotect", files=files)
        response_time = time.time() - start_time
        passed = response.status_code == 400
        return log_test_result("POST /unprotect - 上传非xlsx文件(.xls)", "POST", "/unprotect", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("POST /unprotect - 上传非xlsx文件(.xls)", "POST", "/unprotect", -1, response_time, str(e), False)

def test_unprotect_get_method():
    start_time = time.time()
    try:
        response = requests.get(f"{BASE_URL}/unprotect")
        response_time = time.time() - start_time
        passed = response.status_code == 405
        return log_test_result("GET /unprotect (不支持的方法)", "GET", "/unprotect", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("GET /unprotect (不支持的方法)", "GET", "/unprotect", -1, response_time, str(e), False)

def test_unprotect_put_method():
    start_time = time.time()
    try:
        response = requests.put(f"{BASE_URL}/unprotect")
        response_time = time.time() - start_time
        passed = response.status_code == 405
        return log_test_result("PUT /unprotect (不支持的方法)", "PUT", "/unprotect", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("PUT /unprotect (不支持的方法)", "PUT", "/unprotect", -1, response_time, str(e), False)

def test_unprotect_delete_method():
    start_time = time.time()
    try:
        response = requests.delete(f"{BASE_URL}/unprotect")
        response_time = time.time() - start_time
        passed = response.status_code == 405
        return log_test_result("DELETE /unprotect (不支持的方法)", "DELETE", "/unprotect", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("DELETE /unprotect (不支持的方法)", "DELETE", "/unprotect", -1, response_time, str(e), False)

def test_nonexistent_endpoint():
    start_time = time.time()
    try:
        response = requests.get(f"{BASE_URL}/nonexistent")
        response_time = time.time() - start_time
        passed = response.status_code == 404
        return log_test_result("访问不存在的端点", "GET", "/nonexistent", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("访问不存在的端点", "GET", "/nonexistent", -1, response_time, str(e), False)

def test_unprotect_no_filename():
    start_time = time.time()
    try:
        with open(TEST_FILE_PATH, "rb") as f:
            files = {"file": ("", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            response = requests.post(f"{BASE_URL}/unprotect", files=files)
        response_time = time.time() - start_time
        passed = response.status_code == 400
        return log_test_result("POST /unprotect - 上传无filename的文件", "POST", "/unprotect", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("POST /unprotect - 上传无filename的文件", "POST", "/unprotect", response.status_code if 'response' in locals() else -1, response_time, str(e), False)

def test_unprotect_no_file():
    start_time = time.time()
    try:
        response = requests.post(f"{BASE_URL}/unprotect")
        response_time = time.time() - start_time
        passed = response.status_code == 422
        return log_test_result("POST /unprotect - 无文件上传", "POST", "/unprotect", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("POST /unprotect - 无文件上传", "POST", "/unprotect", -1, response_time, str(e), False)

def test_unprotect_invalid_excel():
    start_time = time.time()
    try:
        files = {"file": ("invalid.xlsx", b"This is not a valid excel file content.", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        response = requests.post(f"{BASE_URL}/unprotect", files=files)
        response_time = time.time() - start_time
        passed = response.status_code == 500
        return log_test_result("POST /unprotect - 无效的Excel文件内容", "POST", "/unprotect", response.status_code, response_time, response.json(), passed)
    except Exception as e:
        response_time = time.time() - start_time
        return log_test_result("POST /unprotect - 无效的Excel文件内容", "POST", "/unprotect", -1, response_time, str(e), False)

def main():
    print(f"\n{'#'*70}")
    print(f"Excel Unprotector API 测试报告")
    print(f"测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'#'*70}")

    test_results = []
    
    test_results.append(test_health_check())
    test_results.append(test_unprotect_valid_file())
    test_results.append(test_unprotect_empty_filename())
    test_results.append(test_unprotect_txt_file())
    test_results.append(test_unprotect_xls_file())
    test_results.append(test_unprotect_get_method())
    test_results.append(test_unprotect_put_method())
    test_results.append(test_unprotect_delete_method())
    test_results.append(test_nonexistent_endpoint())
    test_results.append(test_unprotect_no_filename())
    test_results.append(test_unprotect_no_file())
    test_results.append(test_unprotect_invalid_excel())

    print(f"\n{'#'*70}")
    print(f"测试汇总")
    print(f"{'#'*70}")
    
    passed_count = sum(1 for r in test_results if r["passed"])
    failed_count = len(test_results) - passed_count
    avg_response_time = sum(r["response_time"] for r in test_results) / len(test_results)

    print(f"总测试用例数: {len(test_results)}")
    print(f"通过: {passed_count}")
    print(f"失败: {failed_count}")
    print(f"通过率: {passed_count/len(test_results)*100:.1f}%")
    print(f"平均响应时间: {avg_response_time:.3f}秒")

    if failed_count > 0:
        print(f"\n失败的测试用例:")
        for r in test_results:
            if not r["passed"]:
                print(f"  - {r['test_name']} (状态码: {r['status_code']})")
        return False
    else:
        print(f"\n🎉 所有测试用例均通过！")
        return True

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)