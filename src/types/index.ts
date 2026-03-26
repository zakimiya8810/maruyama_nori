export interface User {
  id: string
  employeeCode: string
  loginId: string
  name: string
  email: string
  role: string
  departmentId: string
  departmentName: string
  division: string
  adminPermission: string
}

export interface Department {
  id: string
  code: string
  name: string
  division: string
  displayOrder: number
}

export interface Employee {
  id: string
  employeeCode: string
  loginId: string
  name: string
  email: string
  role: string
  departmentId: string
  departmentName: string
  division: string
  adminPermission: string
  isRetired: boolean
  displayOrder: number
}

export interface Rank {
  id: string
  code: string
  name: string
  visibility: string
  displayOrder: number
}

export interface BusinessType {
  id: string
  code: string
  name: string
}

export interface PaymentTerm {
  id: string
  pattern: string
  displayOrder: number
}

export interface Customer {
  id: string
  customerCode: string
  name: string
  shortName: string
  kanaName: string
  postalCode: string
  address1: string
  address2: string
  tel: string
  fax: string
  businessTypeId: string
  businessTypeName: string
  rankId: string
  rankName: string
  handlerEmployeeId: string
  handlerName: string
  departmentName: string
  division: string
  billingCustomerId: string
  paymentTermId: string
  paymentTermName: string
  taxMethod: string
  creditLimit: number
  shippingLocation: string
  deliveryMethod: string
  groupCode: string
  representativeName: string
  contactPerson: string
  noriUsage: number
  teaUsage: number
  naturalFoodHandling: string
  shippingFee: string
  registrationDate: string
  tradeStartDate: string
  deleteFlag: number
  billingPostalCode: string
  billingName: string
  billingAddress1: string
  billingAddress2: string
  billingTel: string
  billingFax: string
  lastVisitDate?: string
  needsAlert: boolean
  hasJutaikin: boolean
  isHidden: boolean
}

export interface Product {
  id: string
  productCode: string
  name: string
  category: string
  standardPrice: number
  unit: string
}

export interface CustomerPrice {
  id: string
  customerId: string
  productId: string
  productCode: string
  productName: string
  price: number
  effectiveDate: string
}

export interface Meeting {
  id: string
  customerId: string
  customerCode: string
  customerName: string
  rankName: string
  handlerId: string
  handlerName: string
  departmentName: string
  division: string
  scheduleDate: string
  actualDate: string | null
  appointmentType: string
  purpose: string
  planNotes: string
  result: string
  resultNotes: string
  delayReason: string
  status: 'upcoming' | 'completed' | 'delayed' | 'overdue'
  createdAt: string
}

export interface Application {
  id: string
  applicationId: string
  groupId: string
  applicationType: string
  targetMaster: string
  customerId: string
  customerCode: string
  customerName: string
  applicantId: string
  applicantName: string
  applicantDepartment: string
  applicationDate: string
  status: string
  stage: string
  effectiveDate: string
  reflectionStatus: string
  rejectReason: string
  requiredFields: string[]
  resubmitCount: number
  originalApplicationId: string
}

export interface ApplicationDetail extends Application {
  detailData: ApplicationCustomerData[]
  customerDetailData?: ApplicationCustomerData[]
  prices?: ApplicationPriceData[]
  backupData?: Record<string, unknown>
  priceBackupData?: CustomerPrice[]
  supervisorApprover?: string
  supervisorApprovedAt?: string
  adminApprover?: string
  adminApprovedAt?: string
  directorApprover?: string
  directorApprovedAt?: string
  ceoApprover?: string
  ceoApprovedAt?: string
}

export interface ApplicationCustomerData {
  field: string
  before: string
  after: string
}

export interface ApplicationPriceData {
  productCode: string
  productName: string
  registrationType: string
  priceBefore: number | null
  priceAfter: number | null
}

export interface Jutaikin {
  id: string
  customerId: string
  amount: number
  baseYearMonth: string
}

export interface CustomerSales {
  id: string
  customerId: string
  year: number
  month: number
  amount: number
}

export type ApprovalStage = '申請中' | '上長承認済' | '管理承認済' | '常務承認済' | '決裁完了' | '却下'

export type UserRole = 'general' | 'supervisor' | 'admin' | 'director' | 'ceo'
