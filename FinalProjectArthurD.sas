*
Programmed by: David Arthur
Programmed on: 2021-04-21
Programmed to: Complete Final Project 
;

*Establish librefs and filerefs for incoming files;
x cd "L:\ST555\Data\BookData\Data\Clinical Trial Case Study";
libname InputDS ".";
filename RawData ".";

x cd "L:\ST555\Results";
libname Results ".";

*Establish librefs and filerefs for outgoing files;
x cd "S:\SAS Working Directory\Final Project";
libname Final ".";

*Set options;
options nodate fmtsearch = (Final);

*Open output destinations;
ods pdf file = 'Arthur Final Report Site 1.pdf' dpi = 300 style = Meadow;
ods rtf file = 'Arthur Final Tables Site 1.rtf' style = Sapphire;
ods powerpoint file = 'Arthur Final Presentation Site 1.pptx' style = PowerPointDark;
ods noproctitle;
ods powerpoint exclude all;
ods listing exclude all;

*Create macro variables;
%let visitInput =
  input Subject sf_reason : $20. screen Sex $ DOV : date. notif_date : date. sbp dbp bpu $ 
        pulse pu : $10. Pos : $10. Temperature tu $ Weight wu $ pain;

%let visitLabel =
  label sf_reason = 'Screen Failure Reason'
        notif_date = 'Failure Notification Date'
        DOV = 'Date of Visit'
        screen = 'Screening Flag, 0=Failure, 1=Pass'
        ;

%let visitClean1 =
  bpu = tranwrd(bpu, '-', '/');

%let visitClean2 =
  pos = put(pos, $pos.);

%let labInput =
  input Subject DOV : date. Notif_date : date. SF_Reason : $20. screen Sex $ ALB Alk_Phos 
        ALT AST D_Bili GGTP c_gluc U_Gluc T_Bili Prot hemoglob Hematocr Preg;
%let labLabel =
  label alb = 'Chem-Albumin, g/dL'
        Alk_Phos = 'Chem-Alk. Phos., IU/L'
        ALT = 'Chem-Alt, IU/L'
        AST = 'Chem-AST, IU/L'
        D_Bili = 'Chem-Dir. Bilirubin, mg/dL'
        GGTP = 'Chem-GGTP, IU/L'
        c_gluc = 'Chem-Glucose, mg/dL'
        U_Gluc = 'Uri.-Glucose, 1label =high'
        T_Bili = 'Chem-Tot. Bilirubin, mg/dL'
        Prot = 'Chem-Tot. Prot., g/dL'
        hemoglob = 'Hemoglobin, g/dL'
        Hematocr = 'EVF/PCV, %'
        Preg = 'Pregnancy Flag, 1=Pregnant, 0=Not'
  ;

*Create custom formats;
proc format library = Final;
  value screen
    0 = 'Fail'
    1 = 'Pass'
  ;
  value $sex
    'M' = 'Male'
    'F' = 'Female'
  ;
  value sbp(fuzz = 0)
    low -  120  = 'Acceptable (120 or below)'
    120 <- high = 'High'
  ;
  value visit
    0 = 'Baseline'
    1 = '3 Month'
    2 = '6 Month'
    3 = '9 Month'
    4 = '12 Month'
  ;
  value visitTwo
    0 = 'Baseline Visit'
    1 = '3 Month Visit'
    2 = '6 Month Visit'
    3 = '9 Month Visit'
    4 = '12 Month Visit'
  ;
  value $pos
    'r' -< 's' = 'Recumbent'
    's' -  'z' = 'Seated'
  ;

*Read in raw files;
data baselineVisitSite1;
  infile RawData("Site 1, Baseline Visit.txt") dsd dlm = '09'x missover;
  &visitInput;
  &visitLabel;
run;

data threeMonthVisitSite1;
  infile RawData("Site 1, 3 Month Visit.txt") dsd dlm = '09'x missover;
  &visitInput;
  &visitLabel;
run;


data sixMonthVisitSite1;
  infile RawData("Site 1, 6 Month Visit.txt") dsd dlm = '09'x missover;
  &visitInput;
  &visitLabel;
run;

data nineMonthVisitSite1;
  infile RawData("Site 1, 9 Month Visit.txt") dsd dlm = '09'x missover;
  &visitInput;
  &visitLabel;
run;

data twelveMonthVisitSite1;
  infile RawData("Site 1, 12 Month Visit.txt") dsd dlm = '09'x missover;
  &visitInput;
  &visitLabel;
run;

data baselineLabSite1;
  infile RawData("Site 1, Baseline Lab Results.txt") dsd dlm = '09'x missover;
  &labInput;
  &labLabel;
run;

data threeMonthLabSite1;
  infile RawData("Site 1, 3 Month Lab Results.txt") dsd dlm = '09'x missover;
  &labInput;
  &labLabel;
run;

data sixMonthLabSite1;
  infile RawData("Site 1, 6 Month Lab Results.txt") dsd dlm = '09'x missover;
  &labInput;
  &labLabel;
run;

data nineMonthLabSite1;
  infile RawData("Site 1, 9 Month Lab Results.txt") dsd dlm = '09'x missover;
  &labInput;
  &labLabel;
run;

data twelveMonthLabSite1;
  infile RawData("Site 1, 12 Month Lab Results.txt") dsd dlm = '09'x missover;
  &labInput;
  &labLabel;
run;

*Create frequency tables;
title 'Output 8.2.3: Sex Versus Pain Level Split on Screen Failure Versus Pass, Site 1 Baseline';
proc freq data = baselineVisitSite1;
  table screen*sex*pain / nocol;
  format screen screen.
         sex $sex.
  ;
run;

ods powerpoint exclude none;
title 'Output 8.2.4: Sex Versus Screen Failure at Baseline Visit in Site 1';
proc freq data = baselineVisitSite1(where = (screen = 0));
  table sex*sf_reason / nocol nopercent;
  format sex $sex.;
run;

*Create statistical summary for BP readings;
title 'Output 8.2.5: Diastolic Blood Pressure and Pulse Summary Statistics at Baseline Visit in Site 1';
proc means data = baselineVisitSite1 maxdec = 1 fw = 10;
  class sex sbp;
  var dbp pulse;
  format sex $sex.
         sbp sbp.
  ;
  label sbp = 'Systolic Blood Pressure';
run;

ods powerpoint exclude all;
*Create statistical summary for lab results;
title 'Output 8.2.6: Glucose and Hemoglobin Summary Statistics from Baseline Lab Results, Site 1';
proc means data = baselineLabSite1 min Q1 median Q3 max maxdec=1 nolabels;
  class sex;
  var c_gluc hemoglob;
  format sex $sex.;
  *output out = albumin;
run;

*Create bar charts;
ods rtf exclude all;
ods listing exclude none;
ods listing image_dpi = 300;
ods graphics / reset width = 6in imagename = "8-3-2";
title 'Output 8.3.2: Recruits that Pass Initial Screening, by Month in Site 1';
proc sgplot data = baselineVisitSite1(where = (screen = 1));
  vbar DOV / group = sex
             groupdisplay = cluster;
  format DOV MONYY7.;
  keylegend / location = inside position = topleft title = "";
  xaxis label = 'Month';
  yaxis label = 'Passed Screening at Baseline Visit';
run;

ods listing image_dpi = 300;
ods graphics / reset width = 6in imagename = "8-3-3";
title 'Output 8.3.3: Average Albumin Results-Baseline Lab, Site 1';
proc sgplot data = baselineLabSite1;
  hbar sex / response = alb
             stat = mean
             limits = upper;
  yaxis display = (nolabel);
run;

*Interleave Visit data sets and clean data;
data final.visitsSite1Combined;
  attrib VisitC     length = $20
         VisitNum
         VisitMonth format = best12.
  ;
  set baselineVisitSite1(in = inBaseline)
      threeMonthVisitSite1(in = inThreeMonth)
      sixMonthVisitSite1(in = inSixMonth)
      nineMonthVisitSite1(in = inNineMonth)
      twelveMonthVisitSite1(in = inTwelveMonth)
  ;
  by Subject;
  bpu = tranwrd(bpu, '-', '/');
  pos = put(lowcase(pos), $pos.);
  if(upcase(tu) eq 'C') then do;
    temperature = 1.8 * temperature + 32;
    tu = 'F';
    end;
  if(upcase(wu) eq 'KG') then do;
    weight = 2.20462 * weight;
    wu = 'lb';
    end;
  VisitNum = 0*inBaseline + 1*inThreeMonth + 2*inSixMonth + 3*inNineMonth + 4*inTwelveMonth;
  VisitC = put(VisitNum, visit.);
  VisitMonth = 3*VisitNum;
run;

*Interleave Lab data sets;
data final.LabSite1Combined;
  attrib VisitC     length = $20
         VisitNum
         VisitMonth format = best12.
  ;
  set baselineLabSite1(in = inBaseline)
      threeMonthLabSite1(in = inThreeMonth)
      sixMonthLabSite1(in = inSixMonth)
      nineMonthLabSite1(in = inNineMonth)
      twelveMonthLabSite1(in = inTwelveMonth)
  ;
  by Subject;
  VisitNum = 0*inBaseline + 1*inThreeMonth + 2*inSixMonth + 3*inNineMonth + 4*inTwelveMonth;
  VisitC = put(VisitNum, visit.);
  VisitMonth = 3*VisitNum;
run;

*Create glucose bar charts;
ods listing image_dpi = 300;
ods graphics / reset width = 6in imagename = "8-4-3";
title 'Output 8.4.3: Glucose Distributions, Baseline and 3 Month Visits, Site 1';
proc sgpanel data = final.LabSite1Combined(where = (visitNum in (0,1)));
  panelby visitNum / novarname;
  histogram c_gluc;
  format visitNum visit.;
run;

*Create sbp summary for high-low plot;
ods exclude summary;
proc means data = final.visitsSite1Combined q1 q3;
  class VisitNum;
  var sbp;
  ods output summary = sbpSpan;
run;


*Create sbp bar chart;
ods listing image_dpi = 300;
ods graphics / reset width = 6in imagename = "8-4-4";
title 'Output 8.4.4: Systolic Blood Pressure Quartiles, Site 1';
proc sgplot data = sbpSpan;
  highlow x = VisitNum low = sbp_q1 high = sbp_q3
          / type = bar barwidth = 0.3;
  format VisitNum visit.;
  xaxis label = 'Visit';
  yaxis label = 'Systolic BP--Q1 to Q3 Span'
        values = (95 to 125 by 5);
run;

*Transpose combined visit data set;
proc transpose data = final.visitsSite1Combined(where = (screen=1))
               out = final.visitsTransposed(rename = (_0=vis1 _1=vis2 _2=vis3 _3=vis4 _4=vis5)
                                      drop = _name_ _label_);
  by subject;
  id visitNum;
  var DOV;
run;

*Produce table with Days on Study;
ods listing exclude all;
ods rtf exclude none;
title 'Output 8.5.6: Visits for All Subjects with Days on Study, Site 1 (Partial Listing)';
proc report data = final.visitsTransposed(where = (not missing(vis2)) obs=8) nowd
            style(header) = [textalign = right];
  columns subject--vis5 Days;
  define subject  / display 'subject';
  define vis:     / display format = mmddyy10.;
  define Days     / computed;
  compute Days;
    if not missing(vis5) then Days = vis5 - vis1;
      else if not missing(vis4) then Days = vis4 - vis1;
        else if not missing(vis3) then Days = vis3 - vis1;
          else if not missing(vis2) then Days = vis2 - vis1;
  endcomp;
run;

*Transpose lab data set to match merge with Lab_Info data set;
proc transpose data = final.LabSite1Combined(rename = (alk_phos=ALP c_gluc=C_GLUC d_bili=BILDIR ggtp=GGT hematocr=HCT hemoglob=HGB preg=PREG prot=PROT t_bili=BILI u_gluc=U_GLUC))
               out = final.LabSite1CombinedVert(rename = (_name_=lbtestcd col1=Value ));
  by subject;
  var ALB -- preg;
run;

*Distinguish 2 different glucose tests in Lab_Info data set;
data Lab_Info2;
  set InputDS.Lab_Info;
  if(upcase(lbtestcd) eq 'GLUC') then do;
    if missing(colunits) then lbtestcd = 'U_GLUC';
      else lbtestcd = 'C_GLUC';
  end;
run;

*Sort lab data sets prior to match merge;
proc sort data = final.LabSite1CombinedVert out = final.LabSite1CombinedVertSorted;
  by lbtestcd;
run;

proc sort data = Lab_Info2 out = Lab_Info2Sorted;
  by lbtestcd;
run;

*Match merge lab data set with Lab-Info data set;
data final.labSite1Merged;
  length lbtestcd $ 200;
  merge final.LabSite1CombinedVertSorted Lab_Info2Sorted;
  by lbtestcd;
  if((value ge lownorm and value le highnorm) or nmiss(value, lownorm, highnorm) ne 0) then RangeFlag = 0;
    else RangeFlag = 1;
run;

*Produce lab results table for subject 100;
title 'Output 8.5.8: Baseline Lab Results for Subject 100, Site 1--Including Flag for Values Outside Normal Range';
proc report data = final.labSite1Merged(where = (subject = 100)) nowd;
  columns subject labtest colunits value lownorm highnorm RangeFlag;
  define subject    / display style(header) = [textalign = right];
  define labtest    / display style(header) = [textalign = left]
                      style(column) = [cellwidth = 1.5in];
  define colunits   / display style(header) = [textalign = left];
  define value      / display style(header) = [textalign = right]
                      format = 6.2;
  define lowNorm    / display style(header) = [textalign = right]
                      format = 5.1;
  define highNorm   / display style(header) = [textalign = right]
                      format = 5.1;
  define rangeFlag  / display style(header) = [textalign = right];
run;

*Rotate Visits data set with arrays;
data final.VisitsSite1Rotated(drop = Obs) final.BaselineSite1Rotated;
  set final.visitsSite1Combined;
  length name $20;
  array names[5] $20 _temporary_;
  names[1] = 'Systolic BP';
  names[2] = 'Diastolic BP';
  names[3] = 'Pulse';
  names[4] = 'Temperature';
  names[5] = 'Weight';
  array msmt[*] sbp dbp pulse temperature weight;
  array unit[*] $ bpu bpu pu tu wu;
  do msmtNum = 1 to dim(msmt);
    name = names[msmtNum];
    value = msmt[msmtNum];
    units = unit[msmtNum];
    if(visitNum eq 0) then do;
      Obs +1;
      output final.BaselineSite1Rotated;
      end;
    output final.VisitsSite1Rotated;
  end;
  drop sbp dbp pulse temperature weight bpu pu tu wu;
run;

*Produce table with rotated data;
title 'Output 8.6.1: Rotated Data from Baseline Visit, Site 1 (Partial Listing)';
proc report data = final.BaselineSite1Rotated(obs = 10) nowd
  style(header) = [backgroundcolor = lightblue];
  columns obs subject dov name value units;
  define obs      / display style(column) = [backgroundcolor = lightblue]
                    style(header) = [textalign = right];
  define subject  / display 'subject'
                    style(header) = [textalign = right];
  define dov      / display format = mmddyy10. 'Date of Visit'
                    style(header) = [textalign = right]
                    style(column) = [cellwidth = 1.1in];
  define name     / display
                    style(header) = [textalign = left];
  define value    / display format = 6.1
                    style(header) = [textalign = right];
  define units    / display format = $8.
                    style(header) = [textalign = left];
  attrib dov   format = mmddyy10. label = 'Date of Visit'
         value format = 6.1
         units format = $8.;
run;

*Produce summary report table and output data set to be used for 8.7.3;
title 'Output 8.6.2: Summary Report on Selected Vital Signs, All Visits, Site 1';
proc report data = final.VisitsSite1Rotated(where = (name ne 'Weight')) nowd
            out = final.summaryReport;
  columns visitNum name units value=meanValue value=medianValue value=stdValue value=minValue value=maxValue;
  define visitNum     / group 'Visit' order = internal format = visitTwo.
                        style(column) = [cellwidth = .75in textalign = left];
  *Used SAS Documentation to remember how to order by internal value rather than by formatted value;
  *https://documentation.sas.com/doc/en/pgmsascdc/9.4_3.5/proc/p0wy1vqwvz43uhn1g77eb5xlvzqh.htm#p0vk3hzz34jjugn16qvlcc1bfuv4;
  define name         / group 'Measurement';
  define units        / group 'Units' format = $8.;
  define meanValue    / analysis mean 'Mean' format = 5.1;
  define medianValue  / analysis median 'Median' format = 5.1;
  define stdValue     / analysis std 'Std. Deviation' format = 5.2;
  define minValue     / analysis min 'Minimum' format = 5.1;
  define maxValue     / analysis max 'Maximum' format = 5.1;
run;

*Produce BP summaries;
ods powerpoint exclude none;
title 'Output 8.6.3: BP Summaries, All Visits, Site 1';
proc report data = final.VisitsSite1Rotated(where = (name in ('Systolic BP', 'Diastolic BP'))) nowd;
  columns visitNum name, value, (mean median std);
  define visitNum / group 'Visit' order = internal format = visitTwo.
                    style(column) = [textalign = left];
  define name     / across 'Measurement';
  define value    / ' ';
  define mean     / 'Mean' format = 5.1;
  define median   / 'Median' format = 5.1;
  define std      / 'Std. Dev.' format = 5.2;
run;

*Read in adverse events data;
data adverseSite1;
  infile RawData("Site 1 Adverse Events.txt") dsd dlm = '09'x missover;
  retain subject;
  format stdt endt date9.;
  input   ind0 $1@;
  Obs + 1;
  select(ind0);
    when('s') do;
      input subject 
          / ind1 $1 stdtChar : $9. endtChar : $9.
          / ind2 $1 aetext : $60.
          / ind3 $1 ptcode soccode lltcode hltcode hlgtcode
          / ind4 $1 llterm : $60. hlterm : $60. hlgterm : $60. prefterm : $60. bodysys : $60.
          / ind5 $1 aerel aesev aeser : $1. aeaction dose;
      end;
    when('d') do;
      input stdtChar : $9. endtChar : $9.
          / ind2 $1 aetext : $60.
          / ind3 $1 ptcode soccode lltcode hltcode hlgtcode
          / ind4 $1 llterm : $60. hlterm : $60. hlgterm : $60. prefterm : $60. bodysys : $60.
          / ind5 $1 aerel aesev aeser : $1. aeaction dose;
      end;
  end;
  stdt = input(stdtChar, date9.);
  endt = input(endtChar, date9.);
run;

*Produce adverse events table;
title 'Output 8.7.1: Adverse Events, Site 1 (Partial Listing)';
proc report data = adverseSite1(keep = Obs subject aetext stdtChar endtChar obs = 10) nowd
            style(header) = [backgroundcolor = lightblue];
  columns Obs subject aetext stdtChar endtChar;
    define Obs      / display
                      style(column) = [backgroundcolor = lightblue]
                      style(header) = [textalign = right];
    define subject  / display
                      style(header) = [textalign = right];
    define aetext   / display
                      style(header) = [textalign = left];
    define stdtChar / display 'stdt'
                      style(header) = [textalign = left];
    define endtChar / display 'endt'
                      style(header) = [textalign = left];
run;

ods powerpoint exclude all;

*Add count variable to summary report data set for color cycling;
data final.summaryReportBy;
  set final.summaryReport;
  by visitNum;
  if(first.visitNum or count eq 4) then call missing(count);
  count + 1;
run;

*Produce enhanced summary report table;
title 'Output 8.7.3: Summary Report on Selected Vital Signs, All Visits, Site 1--Enhanced';
proc report data = final.summaryReportBy nowd
            style(lines) = [color = white]
            style(header) = [textalign = left fontsize = 9pt fontweight = bold];
  columns visitNum name units meanValue medianValue stdValue minValue maxValue count color;
  define visitNum     / order 'Visit' order = internal format = visitTwo.
                        style(column) = [fontweight = bold fontsize = 12pt textalign = left];
  *Used SAS Documentation to remember how to order by internal value rather than by formatted value;
  *https://documentation.sas.com/doc/en/pgmsascdc/9.4_3.5/proc/p0wy1vqwvz43uhn1g77eb5xlvzqh.htm#p0vk3hzz34jjugn16qvlcc1bfuv4;
  define name         / display 'Test';
  define units        / display 'Units'     format = $8.;
  define meanValue    / display 'Mean'      format = 5.1;
  define medianValue  / display 'Median'    format = 5.1;
  define stdValue     / display 'Std. Dev.' format = 5.2;
  define minValue     / display 'Min.'      format = 5.1;
  define maxValue     / display 'Max.'      format = 5.1;
  define count        / display   noprint;
  define color        / computed  noprint;
  compute color;
    color = count;
    if(count eq 1) then call define(_row_, 'style', 'style = [backgroundcolor = cx1b9e77]');
      else if (count eq 2) then call define(_row_, 'style', 'style = [backgroundcolor = cxd95f02]');
        else if (count eq 3) then call define(_row_, 'style', 'style = [backgroundcolor = cx7570b3]');
          else if (count eq 4) then call define(_row_, 'style', 'style = [backgroundcolor = cxe7298a]');
  endcomp;
  compute after visitNum / style = [backgroundcolor = black];
    line "";
  endcomp;
run;

ods rtf close;
ods pdf close;
ods powerpoint close;

quit;
