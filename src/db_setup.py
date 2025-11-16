import os
from sqlalchemy import create_engine, Column, Integer, String, Text, ForeignKey
from sqlalchemy.orm import sessionmaker, declarative_base

# --- 1. 配置資料庫路徑和連線 ---
# 確保資料庫檔案會被創建在我們預期的 /database 資料夾中
BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 取得當前db_setup.py的路徑
DB_PATH = os.path.join(os.path.dirname(BASE_DIR), 'database', 'app.db')
ENGINE_URL = f'sqlite:///{DB_PATH}'

# 創建資料庫引擎
engine = create_engine(ENGINE_URL)

# 聲明一個基礎類，所有ORM類都將繼承它
Base = declarative_base()


# --- 2. 定義 ORM 表格結構 ---

# A. 客戶檔案索引表 (Client_Index)
class ClientIndex(Base):
    __tablename__ = 'client_index'

    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True, nullable=False, comment="客戶名稱 (如: A公司)")
    unified_number = Column(String, unique=True, comment="統一編號/證號")
    current_address = Column(String, comment="公司目前註冊地址")
    legal_rep = Column(String, comment="法定代理人/負責人姓名")

    def __repr__(self):
        return f"<ClientIndex(name='{self.name}')>"


# B. 文件類型需求表 (Doc_Type_Requirements)
class DocTypeRequirement(Base):
    __tablename__ = 'doc_type_requirements'

    id = Column(Integer, primary_key=True)
    business_name = Column(String, unique=True, nullable=False, comment="業務名稱 (如: 地址變更、設立登記)")
    required_docs_json = Column(Text, comment="所需文件清單 (JSON格式字串)")
    notes = Column(Text, comment="業務處理備註說明")

    def __repr__(self):
        return f"<DocTypeRequirement(business_name='{self.business_name}')>"


# C. 範本變數表 (Template_Variables)
class TemplateVariable(Base):
    __tablename__ = 'template_variables'

    id = Column(Integer, primary_key=True)
    template_filename = Column(String, nullable=False, comment="對應的範本檔名 (如: 變更登記表.docx)")
    placeholder_key = Column(String, nullable=False, comment="範本文件中的佔位符 (如: <<COMPANY_NAME>>)")
    variable_source = Column(String, nullable=False,
                             comment="變數的資料來源 (如: ClientIndex.name 或 UserInput.new_address)")

    def __repr__(self):
        return f"<TemplateVariable(template_filename='{self.template_filename}', key='{self.placeholder_key}')>"


# --- 3. 建立表格和連線會話 ---

def setup_database():
    """創建所有的資料庫表格 (如果它們不存在)"""

    # 檢查並建立 /database 資料夾
    db_dir = os.path.join(os.path.dirname(BASE_DIR), 'database')
    if not os.path.exists(db_dir):
        os.makedirs(db_dir)
        print(f"Created directory: {db_dir}")

    # 創建表格
    Base.metadata.create_all(engine)
    print(f"Database setup complete! Tables created in: {DB_PATH}")


# 創建會話類別
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

if __name__ == "__main__":
    setup_database()