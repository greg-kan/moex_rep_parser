--create schema reports;

--drop table reports.brokerage_monthly;

--drop table reports.brokerage_monthly_securities;
--drop table reports.brokerage_monthly_security;

--drop table reports.brokerage_monthly_money;
--drop table reports.brokerage_monthly_operation;

drop table reports.brokerage_monthly_transaction;
drop table reports.brokerage_monthly_transactions;
drop table reports.brokerage_monthly_money_fee;
drop table reports.brokerage_monthly_money_operation;
drop table reports.brokerage_monthly_money;
drop table reports.brokerage_monthly_security;
drop table reports.brokerage_monthly_securities;
drop table reports.brokerage_monthly;

select * from reports.brokerage_monthly bm
  left join reports.brokerage_monthly_securities bs on bs.report_id  = bm.id
  left join reports.brokerage_monthly_security bss on bss.securities_id = bs.id
 where bm."year" = 2024 and bm."month" = 3;

select * from reports.brokerage_monthly bm
  left join reports.brokerage_monthly_money bmm on bmm.report_id  = bm.id
  left join reports.brokerage_monthly_money_operation bmo on bmo.money_id = bmm.id
  
select * from reports.brokerage_monthly bm
  left join reports.brokerage_monthly_money bmm on bmm.report_id  = bm.id
  left join reports.brokerage_monthly_money_fee bmf on bmf.money_id = bmm.id
  
  