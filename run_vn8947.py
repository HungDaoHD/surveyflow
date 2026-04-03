"""Run IngestionStep for VN8947."""
import sys, json, logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s %(message)s')
sys.path.insert(0, r'C:\Users\PC\OneDrive\DevZone\PyPackages\surveyflow')

from surveyflow.steps.ingestion.ingestion_step import IngestionStep

# ── definition ──
definition = {
    "survey": {
        "survey_id": 723122,
        "title": "VN8947 - Khảo sát dân cư khu vực Vũng Tàu.",
        "english_title": "VN8947 - Khảo sát dân cư khu vực Vũng Tàu.",
        "status": "2",
        "start_date": "2026-03-17 15:59:42",
        "end_date": ""
    },
    "questions": [
        {"question_id":795781,"position":1,"question":"Please select the user name","english_question":"Please select the user name who invited you to this survey","type":1,"input_type":73,"mandatory":True,"status":1},
        {"question_id":795780,"position":2,"question":"Record","english_question":"Record","type":40,"input_type":40,"mandatory":True,"status":1},
        {"question_id":795570,"position":3,"question":"Anh/chị tên gì","english_question":"What is your name?","type":1106,"input_type":51,"mandatory":True,"status":1},
        {"question_id":795571,"position":4,"question":"Số điện thoại","english_question":"What is your phone number?","type":1107,"input_type":52,"mandatory":True,"status":1},
        {"question_id":795572,"position":5,"question":"Địa điểm phỏng vấn","english_question":"Input the address","type":1,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795699,"position":6,"question":"Địa chỉ sinh sống","english_question":"Please provide your current address","type":2,"input_type":100,"mandatory":True,"status":1},
        {"question_id":795493,"position":7,"question":"Khoảng cách","english_question":"(Input by fieldworker) Distance from the target store","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795494,"position":8,"question":"Hướng","english_question":"(Input by fieldworker) Direction from the target store","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795495,"position":9,"question":"Người quyết định mua sắm","english_question":"Are you the decision-maker when it comes to purchasing items (such as food, clothing, household appliances, electronics, etc.) for yourself/your family?","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795496,"position":10,"question":"Giới tính","english_question":"Respondent gender","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795603,"position":11,"question":"Tuổi thực","english_question":"How old are you?","type":1101,"input_type":50,"mandatory":True,"status":1},
        {"question_id":795497,"position":12,"question":"Khoảng tuổi","english_question":"Age","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795498,"position":13,"question":"Nghề nghiệp","english_question":"What is your occupation","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795499,"position":14,"question":"Sở hữu ô tô","english_question":"Do you own a car","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795500,"position":15,"question":"Sở hữu xe máy","english_question":"Do you own motorbike(s)?","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795501,"position":16,"question":"Số thành viên gia đình","english_question":"Number of people in a house (including yourself)","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795593,"position":17,"question":"Tình trạng hôn nhân","english_question":"Marital status","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795502,"position":18,"question":"Sống cùng ai","english_question":"Who do you live with","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795503,"position":19,"question":"Tuổi con","english_question":"Children age","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795504,"position":20,"question":"Tuổi chủ hộ","english_question":"Who is the age group of house owner","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795505,"position":21,"question":"Nuôi thú cưng","english_question":"Do you have a pet?","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795506,"position":22,"question":"Loại thú cưng","english_question":"What pets do you have?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795507,"position":23,"question":"Sở hữu nhà","english_question":"Please select the type of house ownership","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795508,"position":24,"question":"Loại nhà","english_question":"What types of housing do you live in?","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795509,"position":25,"question":"Thời gian sinh sống","english_question":"How long do you live in the current housing?","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795510,"position":26,"question":"Báo điện tử","english_question":"What are the news site that you use often?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795511,"position":27,"question":"Thẻ tín dụng","english_question":"What banks of credit cards do you own?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795512,"position":28,"question":"Thu nhập hộ gia đình","english_question":"Average monthly household income (including the amount that has been sent over from someone, or won from the the other business","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795513,"position":29,"question":"Phương tiện đi mua sắm","english_question":"What transportation do you usually use for shopping?","type":5,"input_type":0,"mandatory":True,"status":1},
        {"question_id":796159,"position":30,"question":"Phương thức thanh toán","english_question":"Please select the payment method that you use","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795514,"position":31,"question":"Thanh toán thường xuyên nhất","english_question":"What is the payment method you use the most","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795515,"position":32,"question":"Tần suất mua online","english_question":"How often do you buy online in a month","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795516,"position":33,"question":"Mua online mặt hàng gì","english_question":"What kind of categories do you buy online?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795517,"position":34,"question":"Thời gian rảnh","english_question":"How do you usually spend your free-time","type":5,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795518,"position":35,"question":"Nơi mua thực phẩm top 2","english_question":"Please select top 2 stores that you go to buy foods products","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795519,"position":36,"question":"Tần suất mua thực phẩm","english_question":"How often do you buy foods products in a week?","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795520,"position":37,"question":"Chi tiêu thực phẩm","english_question":"How much do you spend on food per one shopping?","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795574,"position":38,"question":"Lý do chọn nơi mua thực phẩm","english_question":"What are the reasons that you use for the selected 2 stores","type":5,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795521,"position":39,"question":"Ăn ngoài","english_question":"Please select the dishes/restaurants you have visited when eating out.","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795522,"position":40,"question":"Tần suất ăn ngoài","english_question":"How often do you eat the following cuisine outside of your house?","type":4,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795594,"position":41,"question":"Quán ăn thường xuyên","english_question":"Interviewers select dishes/restaurants that respondents had eaten at least once every 2-3 months:","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795523,"position":42,"question":"Chi tiêu ăn ngoài","english_question":"How much do you spend at these restaurant per visit per time?","type":4,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795524,"position":43,"question":"Đi cùng ai","english_question":"Who do you go with?","type":4,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795525,"position":44,"question":"Loại nhà hàng hay đến","english_question":"Please select the restaurants that you go often","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795526,"position":45,"question":"Nhà hàng mong muốn","english_question":"Please share us if there are any restaurant you wish to have around the area","type":1,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795527,"position":46,"question":"Tần suất mua quần áo","english_question":"How often do you buy fashion items?","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795528,"position":47,"question":"Chi tiêu quần áo","english_question":"How much do you pay per one time shopping for fashion items?","type":1,"input_type":3,"mandatory":True,"status":1},
        {"question_id":795529,"position":48,"question":"Kênh mua quần áo","english_question":"What channels do you purchase fashion items at?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795530,"position":49,"question":"Yếu tố mua quần áo","english_question":"What are the things you pay attention to in buying clothes?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795531,"position":50,"question":"Thương hiệu thời trang yêu thích","english_question":"What are your favorite fashion brands","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795532,"position":51,"question":"Thương hiệu thời trang mong muốn","english_question":"Please share us if there are any fashion brand shop you wish to have around the area","type":1,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795533,"position":52,"question":"Tần suất mua quần áo trẻ em","english_question":"How often do you buy kids fashion item","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795534,"position":53,"question":"Chi tiêu quần áo trẻ em","english_question":"How much do you pay per one time shopping for the cateogry?","type":1,"input_type":3,"mandatory":True,"status":1},
        {"question_id":795535,"position":54,"question":"Kênh mua quần áo trẻ em","english_question":"Please share us if there are any specific stores that you buy children fashion items with","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795536,"position":55,"question":"Yếu tố mua quần áo trẻ em","english_question":"What are the things that you pay attention to for the kids fashion items","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795537,"position":56,"question":"Kênh mua đồ chơi","english_question":"What channels do you purchase toy items at?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795538,"position":57,"question":"Yếu tố mua đồ chơi","english_question":"What are the things that you pay attention to for purchasing toys","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795539,"position":58,"question":"Tần suất mua mỹ phẩm","english_question":"How often do you buy beauty items?","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795540,"position":59,"question":"Chi tiêu mỹ phẩm","english_question":"How much do you pay per one time shopping for beauty items?","type":1,"input_type":3,"mandatory":True,"status":1},
        {"question_id":795541,"position":60,"question":"Kênh mua mỹ phẩm","english_question":"What channels do you purchase beauty items at?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795542,"position":61,"question":"Yếu tố mua mỹ phẩm","english_question":"What are the things you pay attention to in buying beauty items","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795543,"position":62,"question":"Sản phẩm mỹ phẩm đang dùng","english_question":"Please select the beauty items that you own","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795544,"position":63,"question":"Thiết bị gia dụng đang dùng","english_question":"Please share us the home appliance / electronics / IT products that you own","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795545,"position":64,"question":"Kênh mua thiết bị gia dụng","english_question":"What channels do you purchase such home appliance / electronics / IT products at?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795546,"position":65,"question":"Yếu tố mua thiết bị gia dụng","english_question":"What are the things you pay attention to in buying electronics / home appliances / IT","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795547,"position":66,"question":"Thương hiệu điện thoại","english_question":"Please select the smartphone brand that you own","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795548,"position":67,"question":"Thương hiệu máy lạnh","english_question":"Please select the Air conditioner brand that you own","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795549,"position":68,"question":"Thương hiệu tủ lạnh","english_question":"Please select the Refrigerator brand that you own","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795550,"position":69,"question":"Thương hiệu máy giặt","english_question":"Please select the Washing machine brand that you own","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795551,"position":70,"question":"Thương hiệu TV","english_question":"Please select the TV brand that you own","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795552,"position":71,"question":"Ngân sách thiết bị gia dụng","english_question":"Please select your budget in case you are to purchase the below","type":4,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795553,"position":72,"question":"Nội thất đang dùng","english_question":"Please share us the furniture that you own","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795554,"position":73,"question":"Kênh mua nội thất","english_question":"What channels do you purchase such furniture at?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795555,"position":74,"question":"Yếu tố mua nội thất","english_question":"What are the things you pay attention to in buying furniture","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795556,"position":75,"question":"Ngân sách nội thất","english_question":"Please select your budget in case you are to purchase the below","type":4,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795557,"position":76,"question":"Hoạt động của con hiện tại","english_question":"What are the things that your children learn","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795558,"position":77,"question":"Hoạt động con tương lai","english_question":"Are there anything that you would like your children to learn in the future?","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795559,"position":78,"question":"Tần suất xem phim","english_question":"How often do you watch movies at cinema?","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795560,"position":79,"question":"Biết TTTM nào","english_question":"Please select the mall that you aware aware in this city","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795561,"position":80,"question":"Tần suất đến TTTM","english_question":"How often do you visit the following mall (in the last 90 days)","type":4,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795562,"position":81,"question":"TTTM thường xuyên nhất","english_question":"Please select the shopping mall you visit the most","type":2,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795563,"position":82,"question":"Lý do đến TTTM đó","english_question":"What are the reasons that you visit  the most  ","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795577,"position":83,"question":"TTTM đến ít nhất 2-3 tháng","english_question":"Interviewer selected supermarkets/shopping malls that respondents had visited at least once in the past 2-3 months.","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795564,"position":84,"question":"Mua gì ở TTTM","english_question":"What do you shop at the selected malls","type":5,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795579,"position":85,"question":"Mức độ hài lòng TTTM","english_question":"How much are you satisfied with the following malls","type":4,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795581,"position":86,"question":"Hài lòng điều gì ở TTTM","english_question":"What are the things that you are satisfied with about these malls","type":5,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795567,"position":87,"question":"Cần cải thiện ở TTTM","english_question":"Please select if there are anything that you wish to improve about the following mall","type":5,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795568,"position":88,"question":"Mong muốn thêm ở TTTM","english_question":"Please share us about the types of stores and services you wish shopping centers to have more facilities or services","type":3,"input_type":0,"mandatory":True,"status":1},
        {"question_id":795575,"position":89,"question":"reward","english_question":"reward","type":1,"input_type":68,"mandatory":True,"status":1},
    ],
    "question_count": 89
}

# ── row pages ──
pages = []
for fpath in [
    r'C:\Users\PC\.claude\projects\C--Users-PC-OneDrive-DevZone-PyPackages-surveyflow--claude-worktrees-clever-boyd\c1b76edf-5bd6-4519-a4cb-f85c372c459f\tool-results\mcp-92b3762a-2e81-47a2-9b3e-095c0b4486d7-get_survey_rows-1775202854470.txt',
    r'C:\Users\PC\.claude\projects\C--Users-PC-OneDrive-DevZone-PyPackages-surveyflow--claude-worktrees-clever-boyd\c1b76edf-5bd6-4519-a4cb-f85c372c459f\tool-results\mcp-92b3762a-2e81-47a2-9b3e-095c0b4486d7-get_survey_rows-1775203012244.txt',
]:
    with open(fpath, encoding='utf-8') as f:
        pages.append(json.load(f))

# page 3 — 1 row (already fetched inline)
pages.append({
    "rows": [{
        "date_time": "2026-03-26 14:35:40", "Key_in_date": "2026-03-26 14:35:40",
        "lastmodified_date": "2026-04-02 10:48:58", "task_id": "426_723122_2968182",
        "profile_status": "approved",
        "questions": [
            {"type":"freetext","question":"Input the address","answer":"160 do luong"},
            {"type":"singlechoice","question":"Please provide your current address","answer":"Ward 11"},
            {"type":"singlechoice","question":"(Input by fieldworker) Distance from the target store","answer":"0 - 2.49 km"},
            {"type":"singlechoice","question":"(Input by fieldworker) Direction from the target store","answer":"East"},
            {"type":"singlechoice","question":"Are you the decision-maker when it comes to purchasing items (such as food, clothing, household appliances, electronics, etc.) for yourself/your family?","answer":"I am the main decision maker for some of the above items"},
            {"type":"singlechoice","question":"Respondent gender","answer":"Female"},
            {"type":"singlechoice","question":"Age","answer":"35-39"},
            {"type":"singlechoice","question":"What is your occupation","answer":"Company Employee"},
            {"type":"singlechoice","question":"Do you own a car","answer":"No"},
            {"type":"singlechoice","question":"Do you own motorbike(s)?","answer":"Yes - Own 2"},
            {"type":"singlechoice","question":"Number of people in a house (including yourself)","answer":"5"},
            {"type":"singlechoice","question":"Marital status","answer":"Married (with children)"},
            {"type":"multiplechoice","question":"Who do you live with","answer":[{"answer_name":"Spouse","answer_id":5548964}]},
            {"type":"singlechoice","question":"Who is the age group of house owner","answer":"Over 60s"},
            {"type":"singlechoice","question":"Do you have a pet?","answer":"No"},
            {"type":"singlechoice","question":"Please select the type of house ownership","answer":"Owned (Parents)"},
            {"type":"singlechoice","question":"What types of housing do you live in?","answer":"Detached House"},
            {"type":"singlechoice","question":"How long do you live in the current housing?","answer":"Over 20 years"},
            {"type":"singlechoice","question":"Average monthly household income (including the amount that has been sent over from someone, or won from the the other business","answer":"25M VND~29.9M VND"},
            {"type":"singlechoice","question":"What is the payment method you use the most","answer":"By cash"},
            {"type":"singlechoice","question":"How often do you buy online in a month","answer":"Less than once/month"},
            {"type":"singlechoice","question":"How often do you buy foods products in a week?","answer":"More than 4 times"},
            {"type":"singlechoice","question":"How much do you spend on food per one shopping?","answer":"101-200K VND"},
            {"type":"singlechoice","question":"How often do you buy fashion items?","answer":"Less than once every 2-3 months"},
            {"type":"singlechoice","question":"How often do you buy kids fashion item","answer":"Less than once every 2-3 months"},
            {"type":"singlechoice","question":"How often do you buy beauty items?","answer":"About 2-3 times a month"},
            {"type":"singlechoice","question":"Please select the shopping mall you visit the most","answer":"Co.opmart Vung Tau"},
            {"type":"singlechoice","question":"How often do you watch movies at cinema?","answer":"Have been only once"},
        ]
    }]
})

total_rows = sum(len(p["rows"]) for p in pages)
print(f"Pages: {len(pages)}, Total rows: {total_rows}")

# ── run ──
step = IngestionStep()
ctx = step.run({
    "definition":     definition,
    "rows_pages":     pages,
    "output_dir":     r"C:\Users\PC\OneDrive\DevZone\PyPackages\surveyflow\output_test\VN8947",
    "profile_status": ["approved"],
})

df   = ctx["rawdata"]
meta = ctx["metadata"]
print(f"\nDataFrame : {df.shape[0]} rows x {df.shape[1]} cols")
print(f"Questions : {len(meta['questions'])} in metadata")
print(f"rawdata   : {ctx['rawdata_path']}")
print(f"metadata  : {ctx['metadata_path']}")

# sample
print("\nSample row (q6, q10, q12):")
print(df[["task_id","q6","q10","q12"]].head(3).to_string(index=False))
