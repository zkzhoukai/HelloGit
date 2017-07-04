' ONLY ONE ARGUMENT TO THIS SCRIPT - THE NETBIOS COMPUTERNAME
set objArgs = Wscript.arguments
if wscript.arguments.count = 0 then
	wscript.quit
end if

'#capture the name of the computer we're running on
NetBiosName = objArgs(0)

Dim intStatus, objHTTP
dim siteArray(280)  'where the URLList will be stored

siteArray(0) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/BulkImportAdjustmentService.svc"
siteArray(1) = "http://whdcepiweb21/WCF/AP1/CRCServices/ExternalINSS/ExternalINSSService.svc"
siteArray(2) = "http://whdcepiweb21/WCF/AP2/CRCServices/ExternalINSS/ExternalINSSService.svc"
siteArray(3) = "http://whdcepiweb21/WCF/AP3/CRCServices/ExternalINSS/ExternalINSSService.svc"
siteArray(4) = "http://whdcepiweb21/WCF/CN2/CRCServices/ExternalINSS/ExternalINSSService.svc"
siteArray(5) = "http://whdcepiweb21/WCF/AP1/CRCServices/TicketDataForwarders/ExternalINSS.svc"
siteArray(6) = "http://whdcepiweb21/WCF/AP2/CRCServices/TicketDataForwarders/ExternalINSS.svc"
siteArray(7) = "http://whdcepiweb21/WCF/AP3/CRCServices/TicketDataForwarders/ExternalINSS.svc"
siteArray(8) = "http://whdcepiweb21/WCF/CN2/CRCServices/TicketDataForwarders/ExternalINSS.svc"
siteArray(9) = "http://whdcepiweb21/WCF/AP1/CRCServices/CDM/CDMService.svc"
siteArray(10) = "http://whdcepiweb21/WCF/AP2/CRCServices/CDM/CDMService.svc"
siteArray(11) = "http://whdcepiweb21/WCF/AP3/CRCServices/CDM/CDMService.svc"
siteArray(12) = "http://whdcepiweb21/WCF/CN2/CRCServices/CDM/CDMService.svc"
siteArray(13) = "http://whdcepiweb21/WCF/AP1/CRCServices/TicketDataForwarders/CDM.svc"
siteArray(14) = "http://whdcepiweb21/WCF/AP2/CRCServices/TicketDataForwarders/CDM.svc"
siteArray(15) = "http://whdcepiweb21/WCF/AP3/CRCServices/TicketDataForwarders/CDM.svc"
siteArray(16) = "http://whdcepiweb21/WCF/CN2/CRCServices/TicketDataForwarders/CDM.svc"
siteArray(17) = "http://whdcepiweb21/WCF/AP1/CRCServices/CDM/CDMInquiryService.svc"
siteArray(18) = "http://whdcepiweb21/WCF/AP2/CRCServices/CDM/CDMInquiryService.svc"
siteArray(19) = "http://whdcepiweb21/WCF/AP3/CRCServices/CDM/CDMInquiryService.svc"
siteArray(20) = "http://whdcepiweb21/WCF/CN2/CRCServices/CDM/CDMInquiryService.svc"
siteArray(21) = "http://whdcepiweb21/WCF/AP1/CRCServices/CustomerServiceService/CustomerServiceService.svc"
siteArray(22) = "http://whdcepiweb21/WCF/AP2/CRCServices/CustomerServiceService/CustomerServiceService.svc"
siteArray(23) = "http://whdcepiweb21/WCF/AP3/CRCServices/CustomerServiceService/CustomerServiceService.svc"
siteArray(24) = "http://whdcepiweb21/WCF/CN2/CRCServices/CustomerServiceService/CustomerServiceService.svc"
siteArray(25) = "http://whdcepiweb21/WCF/AP1/CRCServices/CustomerServiceService/Contest.svc"
siteArray(26) = "http://whdcepiweb21/WCF/AP2/CRCServices/CustomerServiceService/Contest.svc"
siteArray(27) = "http://whdcepiweb21/WCF/AP3/CRCServices/CustomerServiceService/Contest.svc"
siteArray(28) = "http://whdcepiweb21/WCF/CN2/CRCServices/CustomerServiceService/Contest.svc"
siteArray(29) = "http://whdcepiweb21/WCF/AP1/CRCServices/CustomerServiceService/CommissionService.svc"
siteArray(30) = "http://whdcepiweb21/WCF/AP2/CRCServices/CustomerServiceService/CommissionService.svc"
siteArray(31) = "http://whdcepiweb21/WCF/AP3/CRCServices/CustomerServiceService/CommissionService.svc"
siteArray(32) = "http://whdcepiweb21/WCF/CN2/CRCServices/CustomerServiceService/CommissionService.svc"
siteArray(33) = "http://whdcepiweb21/WCF/AP1/CRCServices/CustomerServiceService/CSAdjustmentService.svc"
siteArray(34) = "http://whdcepiweb21/WCF/AP2/CRCServices/CustomerServiceService/CSAdjustmentService.svc"
siteArray(35) = "http://whdcepiweb21/WCF/AP3/CRCServices/CustomerServiceService/CSAdjustmentService.svc"
siteArray(36) = "http://whdcepiweb21/WCF/CN2/CRCServices/CustomerServiceService/CSAdjustmentService.svc"
siteArray(37) = "http://whdcepiweb21/WCF/AP1/CRCServices/CustomerServiceService/CommissionAdjustmentImpact.svc"
siteArray(38) = "http://whdcepiweb21/WCF/AP2/CRCServices/CustomerServiceService/CommissionAdjustmentImpact.svc"
siteArray(39) = "http://whdcepiweb21/WCF/AP3/CRCServices/CustomerServiceService/CommissionAdjustmentImpact.svc"
siteArray(40) = "http://whdcepiweb21/WCF/CN2/CRCServices/CustomerServiceService/CommissionAdjustmentImpact.svc"
siteArray(41) = "http://whdcepiweb21/WCF/AP1/CRCServices/CustomerServiceService/ConsultantFlag.svc"
siteArray(42) = "http://whdcepiweb21/WCF/AP2/CRCServices/CustomerServiceService/ConsultantFlag.svc"
siteArray(43) = "http://whdcepiweb21/WCF/AP3/CRCServices/CustomerServiceService/ConsultantFlag.svc"
siteArray(44) = "http://whdcepiweb21/WCF/CN2/CRCServices/CustomerServiceService/ConsultantFlag.svc"
siteArray(45) = "http://whdcepiweb21/WCF/AP1/CRCServices/CustomerServiceService/OrderProductionTaxRates.svc"
siteArray(46) = "http://whdcepiweb21/WCF/AP2/CRCServices/CustomerServiceService/OrderProductionTaxRates.svc"
siteArray(47) = "http://whdcepiweb21/WCF/AP3/CRCServices/CustomerServiceService/OrderProductionTaxRates.svc"
siteArray(48) = "http://whdcepiweb21/WCF/CN2/CRCServices/CustomerServiceService/OrderProductionTaxRates.svc"
siteArray(49) = "http://whdcepiweb21/WCF/AP1/CRCServices/CustomerServiceService/ProductionService.svc"
siteArray(50) = "http://whdcepiweb21/WCF/AP2/CRCServices/CustomerServiceService/ProductionService.svc"
siteArray(51) = "http://whdcepiweb21/WCF/AP3/CRCServices/CustomerServiceService/ProductionService.svc"
siteArray(52) = "http://whdcepiweb21/WCF/CN2/CRCServices/CustomerServiceService/ProductionService.svc"
siteArray(53) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/SubsidiaryUtilitiesService.svc"
siteArray(54) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/SubsidiaryUtilitiesService.svc"
siteArray(55) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/SubsidiaryUtilitiesService.svc"
siteArray(56) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/SubsidiaryUtilitiesService.svc"
siteArray(57) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/ProductCreditService.svc"
siteArray(58) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/ProductCreditService.svc"
siteArray(59) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/ProductCreditService.svc"
siteArray(60) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/ProductCreditService.svc"
siteArray(61) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/TransactionService.svc"
siteArray(62) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/TransactionService.svc"
siteArray(63) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/TransactionService.svc"
siteArray(64) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/TransactionService.svc"
siteArray(65) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/ConsultantTaxExemptionService.svc"
siteArray(66) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/ConsultantTaxExemptionService.svc"
siteArray(67) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/ConsultantTaxExemptionService.svc"
siteArray(68) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/ConsultantTaxExemptionService.svc"
siteArray(69) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/ConsultantMonthlyDeductionService.svc"
siteArray(70) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/ConsultantMonthlyDeductionService.svc"
siteArray(71) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/ConsultantMonthlyDeductionService.svc"
siteArray(72) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/ConsultantMonthlyDeductionService.svc"
siteArray(73) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/ConsultantTaxCodeService.svc"
siteArray(74) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/ConsultantTaxCodeService.svc"
siteArray(75) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/ConsultantTaxCodeService.svc"
siteArray(76) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/ConsultantTaxCodeService.svc"
siteArray(77) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/CoefficientAmountService.svc"
siteArray(78) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/CoefficientAmountService.svc"
siteArray(79) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/CoefficientAmountService.svc"
siteArray(80) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/CoefficientAmountService.svc"
siteArray(81) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/TransactionSearchService.svc"
siteArray(82) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/TransactionSearchService.svc"
siteArray(83) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/TransactionSearchService.svc"
siteArray(84) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/TransactionSearchService.svc"
siteArray(85) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/PaymentPreferencesService.svc"
siteArray(86) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/PaymentPreferencesService.svc"
siteArray(87) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/PaymentPreferencesService.svc"
siteArray(88) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/PaymentPreferencesService.svc"
siteArray(89) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/InvoiceService.svc"
siteArray(90) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/InvoiceService.svc"
siteArray(91) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/InvoiceService.svc"
siteArray(92) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/InvoiceService.svc"
siteArray(93) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/MonthlyDetailService.svc"
siteArray(94) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/MonthlyDetailService.svc"
siteArray(95) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/MonthlyDetailService.svc"
siteArray(96) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/MonthlyDetailService.svc"
siteArray(97) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/AccountBatchService.svc"
siteArray(98) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/AccountBatchService.svc"
siteArray(99) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/AccountBatchService.svc"
siteArray(100) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/AccountBatchService.svc"
siteArray(101) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/ActOfAcceptanceService.svc"
siteArray(102) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/ActOfAcceptanceService.svc"
siteArray(103) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/ActOfAcceptanceService.svc"
siteArray(104) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/ActOfAcceptanceService.svc"
siteArray(105) = "http://whdcepiweb21/WCF/AP1/CRCServices/ProfileService/Notes.svc"
siteArray(106) = "http://whdcepiweb21/WCF/AP2/CRCServices/ProfileService/Notes.svc"
siteArray(107) = "http://whdcepiweb21/WCF/AP3/CRCServices/ProfileService/Notes.svc"
siteArray(108) = "http://whdcepiweb21/WCF/CN2/CRCServices/ProfileService/Notes.svc"
siteArray(109) = "http://whdcepiweb21/WCF/AP1/CRCServices/ProfileService/CoverPage.svc"
siteArray(110) = "http://whdcepiweb21/WCF/AP2/CRCServices/ProfileService/CoverPage.svc"
siteArray(111) = "http://whdcepiweb21/WCF/AP3/CRCServices/ProfileService/CoverPage.svc"
siteArray(112) = "http://whdcepiweb21/WCF/CN2/CRCServices/ProfileService/CoverPage.svc"
siteArray(113) = "http://whdcepiweb21/WCF/AP1/CRCServices/ProfileService/AAG.svc"
siteArray(114) = "http://whdcepiweb21/WCF/AP2/CRCServices/ProfileService/AAG.svc"
siteArray(115) = "http://whdcepiweb21/WCF/AP3/CRCServices/ProfileService/AAG.svc"
siteArray(116) = "http://whdcepiweb21/WCF/CN2/CRCServices/ProfileService/AAG.svc"
siteArray(117) = "http://whdcepiweb21/WCF/AP1/CRCServices/ProfileService/ProfileUpdate.svc"
siteArray(118) = "http://whdcepiweb21/WCF/AP2/CRCServices/ProfileService/ProfileUpdate.svc"
siteArray(119) = "http://whdcepiweb21/WCF/AP3/CRCServices/ProfileService/ProfileUpdate.svc"
siteArray(120) = "http://whdcepiweb21/WCF/CN2/CRCServices/ProfileService/ProfileUpdate.svc"
siteArray(121) = "http://whdcepiweb21/WCF/AP1/CRCServices/ProfileService/ProfileAuditService.svc"
siteArray(122) = "http://whdcepiweb21/WCF/AP2/CRCServices/ProfileService/ProfileAuditService.svc"
siteArray(123) = "http://whdcepiweb21/WCF/AP3/CRCServices/ProfileService/ProfileAuditService.svc"
siteArray(124) = "http://whdcepiweb21/WCF/CN2/CRCServices/ProfileService/ProfileAuditService.svc"
siteArray(125) = "http://whdcepiweb21/WCF/AP1/CRCServices/ProfileService/ProfileService.svc"
siteArray(126) = "http://whdcepiweb21/WCF/AP2/CRCServices/ProfileService/ProfileService.svc"
siteArray(127) = "http://whdcepiweb21/WCF/AP3/CRCServices/ProfileService/ProfileService.svc"
siteArray(128) = "http://whdcepiweb21/WCF/CN2/CRCServices/ProfileService/ProfileService.svc"
siteArray(129) = "http://whdcepiweb21/WCF/AP1/CRCServices/TicketDataForwarders/ProfileService.svc"
siteArray(130) = "http://whdcepiweb21/WCF/AP2/CRCServices/TicketDataForwarders/ProfileService.svc"
siteArray(131) = "http://whdcepiweb21/WCF/AP3/CRCServices/TicketDataForwarders/ProfileService.svc"
siteArray(132) = "http://whdcepiweb21/WCF/CN2/CRCServices/TicketDataForwarders/ProfileService.svc"
siteArray(133) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/ConsultantCacheService.svc"
siteArray(134) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/ConsultantCacheService.svc"
siteArray(135) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/ConsultantCacheService.svc"
siteArray(136) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/ConsultantCacheService.svc"
siteArray(137) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/DBConsultantCacheService.svc"
siteArray(138) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/DBConsultantCacheService.svc"
siteArray(139) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/DBConsultantCacheService.svc"
siteArray(140) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/DBConsultantCacheService.svc"
siteArray(141) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/WorkItemCacheService.svc"
siteArray(142) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/WorkItemCacheService.svc"
siteArray(143) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/WorkItemCacheService.svc"
siteArray(144) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/WorkItemCacheService.svc"
siteArray(145) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/WorkItemService.svc"
siteArray(146) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/WorkItemService.svc"
siteArray(147) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/WorkItemService.svc"
siteArray(148) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/WorkItemService.svc"
siteArray(149) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/WorkItemService_TDR.svc"
siteArray(150) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/WorkItemService_TDR.svc"
siteArray(151) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/WorkItemService_TDR.svc"
siteArray(152) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/WorkItemService_TDR.svc"
siteArray(153) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/GLDPService.svc"
siteArray(154) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/GLDPService.svc"
siteArray(155) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/GLDPService.svc"
siteArray(156) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/GLDPService.svc"
siteArray(157) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/EmployeeTransactionsService.svc"
siteArray(158) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/EmployeeTransactionsService.svc"
siteArray(159) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/EmployeeTransactionsService.svc"
siteArray(160) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/EmployeeTransactionsService.svc"
siteArray(161) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/RelationshipEvaluationService.svc"
siteArray(162) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/RelationshipEvaluationService.svc"
siteArray(163) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/RelationshipEvaluationService.svc"
siteArray(164) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/RelationshipEvaluationService.svc"
siteArray(165) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/CRCHealthCheckService.svc"
siteArray(166) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/CRCHealthCheckService.svc"
siteArray(167) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/CRCHealthCheckService.svc"
siteArray(168) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/CRCHealthCheckService.svc"
siteArray(169) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/ConsultantGoalService.svc"
siteArray(170) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/ConsultantGoalService.svc"
siteArray(171) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/ConsultantGoalService.svc"
siteArray(172) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/ConsultantGoalService.svc"
siteArray(173) = "http://whdcepiweb21/WCF/AP1/CRCServices/TicketDataForwarders/ConsultantFlagService.svc"
siteArray(174) = "http://whdcepiweb21/WCF/AP2/CRCServices/TicketDataForwarders/ConsultantFlagService.svc"
siteArray(175) = "http://whdcepiweb21/WCF/AP3/CRCServices/TicketDataForwarders/ConsultantFlagService.svc"
siteArray(176) = "http://whdcepiweb21/WCF/CN2/CRCServices/TicketDataForwarders/ConsultantFlagService.svc"
siteArray(177) = "http://whdcepiweb21/WCF/AP1/CRCServices/TicketDataForwarders/ProductReturnService.svc"
siteArray(178) = "http://whdcepiweb21/WCF/AP2/CRCServices/TicketDataForwarders/ProductReturnService.svc"
siteArray(179) = "http://whdcepiweb21/WCF/AP3/CRCServices/TicketDataForwarders/ProductReturnService.svc"
siteArray(180) = "http://whdcepiweb21/WCF/CN2/CRCServices/TicketDataForwarders/ProductReturnService.svc"
siteArray(181) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/MaintainUnitNameService.svc"
siteArray(182) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/MaintainUnitNameService.svc"
siteArray(183) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/MaintainUnitNameService.svc"
siteArray(184) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/MaintainUnitNameService.svc"
siteArray(185) = "http://whdcepiweb21/WCF/AP1/CRCServices/TicketDataForwarders/MaintainUnitNameService.svc"
siteArray(186) = "http://whdcepiweb21/WCF/AP2/CRCServices/TicketDataForwarders/MaintainUnitNameService.svc"
siteArray(187) = "http://whdcepiweb21/WCF/AP3/CRCServices/TicketDataForwarders/MaintainUnitNameService.svc"
siteArray(188) = "http://whdcepiweb21/WCF/CN2/CRCServices/TicketDataForwarders/MaintainUnitNameService.svc"
siteArray(189) = "http://whdcepiweb21/WCF/AP1/CRCServices/TicketDataForwarders/LetterOfIntentService.svc"
siteArray(190) = "http://whdcepiweb21/WCF/AP2/CRCServices/TicketDataForwarders/LetterOfIntentService.svc"
siteArray(191) = "http://whdcepiweb21/WCF/AP3/CRCServices/TicketDataForwarders/LetterOfIntentService.svc"
siteArray(192) = "http://whdcepiweb21/WCF/CN2/CRCServices/TicketDataForwarders/LetterOfIntentService.svc"
siteArray(193) = "http://whdcepiweb21/WCF/AP1/CRCServices/TicketDataForwarders/CarProgramPenaltyService.svc"
siteArray(194) = "http://whdcepiweb21/WCF/AP2/CRCServices/TicketDataForwarders/CarProgramPenaltyService.svc"
siteArray(195) = "http://whdcepiweb21/WCF/AP3/CRCServices/TicketDataForwarders/CarProgramPenaltyService.svc"
siteArray(196) = "http://whdcepiweb21/WCF/CN2/CRCServices/TicketDataForwarders/CarProgramPenaltyService.svc"
siteArray(197) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/TaskCenterConfigurationService.svc"
siteArray(198) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/TaskCenterConfigurationService.svc"
siteArray(199) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/TaskCenterConfigurationService.svc"
siteArray(200) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/TaskCenterConfigurationService.svc"
siteArray(201) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/LetterOfIntentRequestService.svc"
siteArray(202) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/LetterOfIntentRequestService.svc"
siteArray(203) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/LetterOfIntentRequestService.svc"
siteArray(204) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/LetterOfIntentRequestService.svc"
siteArray(205) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/ConsultantTerminationService.svc"
siteArray(206) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/ConsultantTerminationService.svc"
siteArray(207) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/ConsultantTerminationService.svc"
siteArray(208) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/ConsultantTerminationService.svc"
siteArray(209) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/CurrencyService.svc"
siteArray(210) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/CurrencyService.svc"
siteArray(211) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/CurrencyService.svc"
siteArray(212) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/CurrencyService.svc"
siteArray(213) = "http://whdcepiweb21/WCF/AP1/CRCServices/ProfileService/TrusteeProfileUpdate.svc"
siteArray(214) = "http://whdcepiweb21/WCF/AP2/CRCServices/ProfileService/TrusteeProfileUpdate.svc"
siteArray(215) = "http://whdcepiweb21/WCF/AP3/CRCServices/ProfileService/TrusteeProfileUpdate.svc"
siteArray(216) = "http://whdcepiweb21/WCF/CN2/CRCServices/ProfileService/TrusteeProfileUpdate.svc"
siteArray(217) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/TrusteeService.svc"
siteArray(218) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/TrusteeService.svc"
siteArray(219) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/TrusteeService.svc"
siteArray(220) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/TrusteeService.svc"
siteArray(221) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/TrusteeTicketService.svc"
siteArray(222) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/TrusteeTicketService.svc"
siteArray(223) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/TrusteeTicketService.svc"
siteArray(224) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/TrusteeTicketService.svc"
siteArray(225) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/BankingService.svc"
siteArray(226) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/BankingService.svc"
siteArray(227) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/BankingService.svc"
siteArray(228) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/BankingService.svc"
siteArray(229) = "http://whdcepiweb21/WCF/AP1/CRCServices/CAM/ConsultantCompanyService.svc"
siteArray(230) = "http://whdcepiweb21/WCF/AP2/CRCServices/CAM/ConsultantCompanyService.svc"
siteArray(231) = "http://whdcepiweb21/WCF/AP3/CRCServices/CAM/ConsultantCompanyService.svc"
siteArray(232) = "http://whdcepiweb21/WCF/CN2/CRCServices/CAM/ConsultantCompanyService.svc"
siteArray(233) = "http://whdcepiweb21/WCF/AP1/CRCServices/CustomerServiceService/SalesViolationService.svc"
siteArray(234) = "http://whdcepiweb21/WCF/AP2/CRCServices/CustomerServiceService/SalesViolationService.svc"
siteArray(235) = "http://whdcepiweb21/WCF/AP3/CRCServices/CustomerServiceService/SalesViolationService.svc"
siteArray(236) = "http://whdcepiweb21/WCF/CN2/CRCServices/CustomerServiceService/SalesViolationService.svc"
siteArray(237) = "http://whdcepiweb21/WCF/AP1/CRCServices/Imports/ImportService.svc"
siteArray(238) = "http://whdcepiweb21/WCF/AP2/CRCServices/Imports/ImportService.svc"
siteArray(239) = "http://whdcepiweb21/WCF/AP3/CRCServices/Imports/ImportService.svc"
siteArray(240) = "http://whdcepiweb21/WCF/CN2/CRCServices/Imports/ImportService.svc"
siteArray(241) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/TimeZoneService.svc"
siteArray(242) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/TimeZoneService.svc"
siteArray(243) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/TimeZoneService.svc"
siteArray(244) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/TimeZoneService.svc"
siteArray(245) = "http://whdcepiweb21/WCF/AP1/CRCServices/CustomerServiceService/CarProgramService.svc"
siteArray(246) = "http://whdcepiweb21/WCF/AP2/CRCServices/CustomerServiceService/CarProgramService.svc"
siteArray(247) = "http://whdcepiweb21/WCF/AP3/CRCServices/CustomerServiceService/CarProgramService.svc"
siteArray(248) = "http://whdcepiweb21/WCF/CN2/CRCServices/CustomerServiceService/CarProgramService.svc"
siteArray(249) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/MaintenanceService.svc"
siteArray(250) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/MaintenanceService.svc"
siteArray(251) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/MaintenanceService.svc"
siteArray(252) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/MaintenanceService.svc"
siteArray(253) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/ConfigurationService.svc"
siteArray(254) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/ConfigurationService.svc"
siteArray(255) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/ConfigurationService.svc"
siteArray(256) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/ConfigurationService.svc"
siteArray(257) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/ConsultantService.svc"
siteArray(258) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/ConsultantService.svc"
siteArray(259) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/ConsultantService.svc"
siteArray(260) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/ConsultantService.svc"
siteArray(261) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/MessageService.svc"
siteArray(262) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/MessageService.svc"
siteArray(263) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/MessageService.svc"
siteArray(264) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/MessageService.svc"
siteArray(265) = "http://whdcepiweb21/WCF/AP1/CRCServices/TicketDataForwarders/RecruiterAddressValidationService.svc"
siteArray(266) = "http://whdcepiweb21/WCF/AP2/CRCServices/TicketDataForwarders/RecruiterAddressValidationService.svc"
siteArray(267) = "http://whdcepiweb21/WCF/AP3/CRCServices/TicketDataForwarders/RecruiterAddressValidationService.svc"
siteArray(268) = "http://whdcepiweb21/WCF/CN2/CRCServices/TicketDataForwarders/RecruiterAddressValidationService.svc"
siteArray(269) = "http://whdcepiweb21/WCF/AP1/CRCServices/TicketDataForwarders/PendingAdjustmentsService.svc"
siteArray(270) = "http://whdcepiweb21/WCF/AP2/CRCServices/TicketDataForwarders/PendingAdjustmentsService.svc"
siteArray(271) = "http://whdcepiweb21/WCF/AP3/CRCServices/TicketDataForwarders/PendingAdjustmentsService.svc"
siteArray(272) = "http://whdcepiweb21/WCF/CN2/CRCServices/TicketDataForwarders/PendingAdjustmentsService.svc"
siteArray(273) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/ClassParticipationService.svc"
siteArray(274) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/ClassParticipationService.svc"
siteArray(275) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/ClassParticipationService.svc"
siteArray(276) = "http://whdcepiweb21/WCF/CN2/CRCServices/Services/ClassParticipationService.svc"
siteArray(277) = "http://whdcepiweb21/WCF/AP1/CRCServices/Services/BulkImportAdjustmentService.svc"
siteArray(278) = "http://whdcepiweb21/WCF/AP2/CRCServices/Services/BulkImportAdjustmentService.svc"
siteArray(279) = "http://whdcepiweb21/WCF/AP3/CRCServices/Services/BulkImportAdjustmentService.svc"



'#################
'Now read through the array looking for our computer name in the list.  If we find it, do the URL check.
'If we don't find our computer, don't do anything- those are for other computers
'#################
'Prep the HTTP Object First
Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
'create the variable to hold failures
FailedChecks=""
for i = 0 to ubound(siteArray)
	if instr(lcase(siteArray(i)),lcase(netbiosname))<> 0 then 'we found one, check it!
		wscript.echo siteArray(i)
		wscript.sleep 100


		myWebsite = siteArray(i)

    		

    		objHTTP.Open "GET", siteArray(i), False
    		objHTTP.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MyApp 1.0; Windows NT 5.1)"

    		'On Error Resume Next

    		objHTTP.Send
    		intStatus = objHTTP.Status

    		'On Error Goto 0

    		If intStatus = 200 Then
        		PingSite = "OK"
    		Else
        		PingSite = "BAD"
			if len(FailedChecks) = 0 then 'first entry
				FailedChecks = siteArray(i)
			else 'we have more than one failed check but team only wants a single alert with all of them listed
				FailedChecks = FailedChecks & vbcrlf & siteArray(i)
			end if
    		End If
		intstatus = 0
    		'wscript.echo err.number
    		'wscript.echo objhttp.status
    		'wscript.echo PingSite
	end if
next
'################
'DONE WITH CHECKING - NOW DO THE MONITOR STUFF
'################
'################
'CREATE THE STATUS BASED ON LENGTH OF VARIABLE FailedChecks
'################
if len(FailedChecks) = 0 then
	'we are ok...
	CheckStatus = "OK"
else
	CheckStatus = "BAD"
end if
'wscript.echo FailedChecks
Set objHTTP = Nothing
Dim oAPI, oBag
Set oAPI = CreateObject("MOM.ScriptAPI")
Set oBag = oAPI.CreatePropertyBag()
Call oBag.AddValue("Status", CheckStatus)
call OBag.AddValue("FailedChecks",FailedChecks)
Call oAPI.Return(oBag)
set OBag = nothing
set OAPI = nothing