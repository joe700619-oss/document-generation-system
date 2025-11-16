import os
from docx import Document
from datetime import datetime

import json
from sqlalchemy.orm import sessionmaker
from db_setup import engine, SessionLocal
from db_setup import ClientIndex, DocTypeRequirement, TemplateVariable


# --- 1. 定義輔助函式 ---

def add_sample_data(session):
    """
    添加初始範例數據到三個表格中 (僅當數據不存在時才添加)。
    """

    print("\n--- 1. 檢查並添加客戶範例數據 ---")

    # 檢查客戶是否已存在
    if session.query(ClientIndex).filter_by(unified_number="12345678").first() is None:
        client_a = ClientIndex(
            name="A 科技股份有限公司",
            unified_number="12345678",
            current_address="臺北市信義區忠孝東路 100 號 5 樓",
            legal_rep="王小明"
        )
        session.add(client_a)
        print("✅ 客戶 A 科技股份有限公司已添加到會話。")
    else:
        print("👉 客戶 A 科技股份有限公司 (12345678) 已存在，跳過添加。")

    print("\n--- 2. 檢查並添加業務需求範例數據 ---")
    if session.query(DocTypeRequirement).filter_by(business_name="地址變更").first() is None:
        doc_req_address_change = DocTypeRequirement(
            business_name="地址變更",
            required_docs_json=json.dumps([
                "變更登記表",
                "股東會議紀錄/董事會議紀錄",
                "委託書"
            ]),
            notes="需填寫新舊地址資訊，並確認會議紀錄日期。"
        )
        session.add(doc_req_address_change)
        print("✅ 業務需求 '地址變更' 已添加到會話。")
    else:
        print("👉 業務需求 '地址變更' 已存在，跳過添加。")

    print("\n--- 3. 檢查並添加範本變數範例數據 ---")
    # 這裡我們只檢查一個關鍵變數是否存在即可
    if session.query(TemplateVariable).filter_by(placeholder_key="<<COMPANY_NAME>>").first() is None:
        template_var_list = [
            TemplateVariable(
                template_filename="變更登記表.docx",
                placeholder_key="<<COMPANY_NAME>>",
                variable_source="ClientIndex.name"
            ),
            TemplateVariable(
                template_filename="變更登記表.docx",
                placeholder_key="<<OLD_ADDRESS>>",
                variable_source="ClientIndex.current_address"
            ),
            TemplateVariable(
                template_filename="變更登記表.docx",
                placeholder_key="<<NEW_ADDRESS>>",
                variable_source="UserInput.new_address"
            ),
            TemplateVariable(
                template_filename="變更登記表.docx",
                placeholder_key="<<LEGAL_REP>>",
                variable_source="ClientIndex.legal_rep"  # 來源於客戶資料表
            ),
        ]
        session.add_all(template_var_list)
        print("✅ 範本變數已添加到會話。")
    else:
        print("👉 範本變數已存在，跳過添加。")

    try:
        session.commit()
        print("✅ 所有新數據添加成功並已提交。")
    except Exception as e:
        session.rollback()
        # 注意：如果跳過添加後，還是因為其他意外錯誤導致提交失敗，則印出。
        print(f"❌ 數據提交失敗: {e}")


def query_and_display_data(session):
    """
    查詢並顯示剛剛添加的數據。
    """
    print("\n====================================")
    print("✅ 查詢驗證結果：")
    print("====================================")

    # 查詢客戶
    client = session.query(ClientIndex).filter_by(name="A 科技股份有限公司").first()
    print(f"【客戶名稱】: {client.name}, 統一編號: {client.unified_number}")

    # 查詢業務需求
    req = session.query(DocTypeRequirement).filter_by(business_name="地址變更").first()
    required_docs = json.loads(req.required_docs_json)
    print(f"【業務需求】: {req.business_name} 需要文件: {', '.join(required_docs)}")

    # 查詢範本變數
    vars_list = session.query(TemplateVariable).filter_by(template_filename="變更登記表.docx").all()
    print(f"【範本變數】: 變更登記表所需變數 ({len(vars_list)} 個):")
    for var in vars_list:
        print(f"  -> 佔位符: {var.placeholder_key:<20} 來源: {var.variable_source}")



# ... (在 add_sample_data 和 query_and_display_data 之後新增)

def generate_document(session, client_name, business_name, user_input_data):
    """
    根據使用者輸入和資料庫資訊，生成文件。

    Args:
        session: SQLAlchemy 資料庫會話。
        client_name (str): 客戶名稱。
        business_name (str): 業務類型名稱 (如: 地址變更)。
        user_input_data (dict): 使用者輸入的變數 (如: {'NEW_ADDRESS': '新地址'})。
    """
    print(f"\n--- 開始生成 {client_name} 的 {business_name} 文件 ---")

    # 1. 獲取客戶資訊
    client = session.query(ClientIndex).filter_by(name=client_name).first()
    if not client:
        print(f"❌ 找不到客戶：{client_name}")
        return

    # 2. 獲取文件範本變數列表 (我們在此假設所有地址變更都使用 '變更登記表.docx')
    template_filename = "變更登記表.docx"
    template_vars = session.query(TemplateVariable).filter_by(template_filename=template_filename).all()

    # 3. 準備所有替換數據
    data_map = {}

    for var in template_vars:
        key = var.placeholder_key.strip('<>').upper()  # 提取 KEY (如: COMPANY_NAME)

        if var.variable_source.startswith("ClientIndex"):
            # 數據來自客戶資料表
            attr_name = var.variable_source.split('.')[-1]
            data_map[var.placeholder_key] = getattr(client, attr_name)

        elif var.variable_source.startswith("UserInput"):
            # 數據來自使用者輸入
            input_key = var.variable_source.split('.')[-1].upper()
            data_map[var.placeholder_key] = user_input_data.get(input_key, f"[缺少輸入: {input_key}]")

        # 處理 OLD_ADDRESS (我們需要從 client 中獲取舊地址)
        if var.placeholder_key == "<<OLD_ADDRESS>>":
            data_map[var.placeholder_key] = client.current_address  # 舊地址就是客戶當前的地址

        # 處理負責人 (我們需要負責人資訊)
        if var.placeholder_key == "<<LEGAL_REP>>":
            data_map[var.placeholder_key] = client.legal_rep

    # 4. 執行 Word 範本替換
    try:
        # 載入範本
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        template_path = os.path.join(base_dir, 'templates', template_filename)
        document = Document(template_path)

        # 遍歷段落進行替換
        for p in document.paragraphs:
            for old_key, new_value in data_map.items():
                if old_key in p.text:
                    p.text = p.text.replace(old_key, str(new_value))

        # 5. 儲存新文件到客戶資料夾
        output_dir = os.path.join(base_dir, 'clients', client_name)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{client_name}_{business_name}_{timestamp}.docx"
        output_path = os.path.join(output_dir, output_filename)

        document.save(output_path)
        print(f"✅ 文件生成成功！已儲存至：{output_path}")

    except Exception as e:
        print(f"❌ 文件生成失敗：{e}")
        print("請確認 /templates/變更登記表.docx 檔案是否存在，且未被開啟占用。")


# --- 4. 修改主程式入口 (main.py 的 if __name__ == "__main__": 區塊) ---

if __name__ == "__main__":
    with SessionLocal() as session:
        # 1. 添加數據 (確保範例數據存在)
        add_sample_data(session)

        # 2. 查詢數據 (可選，用於確認)
        # query_and_display_data(session)

        # 3. 運行文件生成邏輯
        # 模擬 AI 接收到指令後，傳遞的結構化資料
        user_input = {
            "NEW_ADDRESS": "臺中市西屯區朝馬路 88 號 12 樓"  # 這是 AI 從使用者輸入中提取的新地址
        }

        generate_document(
            session=session,
            client_name="A 科技股份有限公司",
            business_name="地址變更",
            user_input_data=user_input
        )

    print("\n主程式執行完畢。")