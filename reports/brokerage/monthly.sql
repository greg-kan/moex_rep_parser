--create schema reports;

--drop table reports.brokerage_monthly;
--drop table reports.brokerage_monthly_securities;
--drop table reports.brokerage_monthly_security;

select * from reports.brokerage_monthly bm
  left join reports.brokerage_monthly_securities bs on bs.report_id  = bm.id
  left join reports.brokerage_monthly_security bss on bss.securities_id = bs.id
 where bm."year" = 2024 and bm."month" = 3;
 