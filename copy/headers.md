# headers.py

```python

class MasterHeader:
    ID = 'ID'
    CIN11Degital = 'CIN (11 Degital)'
    CIN = 'CIN'
    RealID = 'Real ID'
    ConnectoID = 'Connecto ID'
    ConnectoCreationDate = 'Connecto Creation Date'
    CustomerName = 'Customer Name'
    BookingCountry = 'BookingCountry'
    BookingEntity = 'BookingEntity'
    LocalSrcSysCode = 'LocalSrcSysCode'
    LOB = 'LOB'
    CustomerType = 'Customer Type'
    OBS_FCCRAM_TYP = 'OBS_FCCRAM_TYP'
    CustomerSegment = 'Customer Segment'
    ReviewType = 'Review Type'
    Triggercode = 'Trigger code'
    Triggerdescription = 'Trigger description'
    RiskRating = 'Risk Rating'
    CRTAppian = 'CRT/Appian'
    BatchRemark = 'Batch Remark'
    RequestDate = 'Request Date'
    CaseUpload = 'Case Upload'
    CRTInitiatedYN = 'CRT Initiated (Y/N)'
    CRTInitiateddate = 'CRT Initiated date'
    LastReviewDate = 'Last Review Date'
    OriginalT0 = 'Original T0'
    ActualT0Amended = 'Actual T0 (Amended)'
    T60T45 = 'T60/T45'
    T90T60 = 'T90/T60'
    DueDate = 'Due Date'
    Task = 'Task'
    TaskStatus = 'Task Status'
    TASKAGEINDECIMAL = 'TASK AGE IN DECIMAL'
    T60T45ReminderLTR = 'T60/T45 Reminder LTR'
    T60LTRDATE = 'T60 LTR DATE'
    T90T60WBHLTR = 'T90/T60 WBH LTR'
    T90LTRDate = 'T90 LTR Date'
    WithdrawRequest = 'Withdraw Request'
    ActionforWithdrawRequest = 'Action for Withdraw Request'
    ApprovalCancelDate = 'Approval/Cancel Date'
    CDSCode = 'CDS Code'
    CancelRecompleteWBHComment = 'Cancel/Re-complete WBH Comment'
    RMName = 'RM Name'
    RMStaffID = 'RM Staff ID'
    AssignDate = 'Assign Date'
    StaffID = 'Staff ID'
    CM = 'CM'
    TL = 'TL'
    Comment = 'Comment'
    ActualT0Connecto = 'Actual T0 (Connecto)'
    UrgentCase = 'Urgent Case'
    RetriggerReason = 'Retrigger Reason'
    CompletionLTRSend = 'Completion LTR Send'
    CompletionLTRDATE = 'Completion LTR DATE'
    CaseAge = 'Case Age'
    PUInfoSys = 'PU Info(Sys)'
    PUName = 'PU Name'
    PUphone = 'PU phone'
    PUemail = 'PU email'
    Escalation = 'Escalation'
    EscalationDetail = 'Escalation Detail'
    Remark = 'Remark'
    ITissue = 'IT issue'
    ActualT0 = 'Actual T0'
    T10 = 'T10'
    T10Call = 'T10 Call'
    T30 = 'T30'
    T30Call = 'T30 Call'
    T60T45Mark = 'T60/T45 Mark'
    T80T50 = 'T80/T50'
    T80T50Call = 'T80/T50 Call'
    T90T60Mark = 'T90/T60 Mark'
    CaseStatus = 'Case Status'
    Queue = 'Queue'
    CustomerEngagementLevel = 'Customer Engagement Level'
    PendingCOI = 'Pending COI'
    PendingBusinessAddressProof = 'Pending Business Address Proof'
    PendingID = 'Pending ID'
    PendingKYCInfo = 'Pending KYC Info'
    CustomerNumber = 'Customer Number'

class BowHeader:
    """任务表的列名配置"""
    task_id = "Customer Number"  # 案件编号
    customer_id = "Customer Number"  # 客户名称
    due_date = "Anniversary Date"  # 截止日期
    effort = "Workload Units"  # 工作量权重
    market = "Market"
    legal_entity = "Legal Entity"
    ppm_cdd_source = "PPM CDD Source"
    CNAndLE = "CN & LE"
    risk_rating = "Risk Pyramid Risk Rating"
    LegalEntity = "Legal Entity"
    Market = "Market"
    LineOfBusiness = "Line of Business"

class RamHeader:
    ReviewID = "Real customer Id"
    CIN = "CIN"
    CustomerNumber = "Customer Number"
    CustomerName = "Customer Name"
    LOB = "LOB"
    LineOfBusiness = "Line of Business"
    NextCDDReviewDate = "Next CDD Review Date"
    DueDate = "Anniversary Date"
```
