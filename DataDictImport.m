function DataDictImport(argin)

if nargin<1
    wizardprocess;
else
    if iscellstr(argin)%Cell
        for i=1:numel(argin)
            files=[files;fromstr2xlsfiles(argin{i})];
        end
    elseif isstr(argin)%
        if isdir(argin)
            multifile_process(list_xls_under_path(argin),argin);
        else
            xlsfiles=fromstr2xlsfiles(argin);
%             hwtbar=waitbar(0,'Importing data dictionary');
            for i=1:numel(xlsfiles)
%                 waitbar(i/numel(xlsfiles),hwtbar,sprintf('Reading "%s"\n',xlsfiles{i}));
                read_DD_xls(xlsfiles{i});
            end
%             close(hwtbar);
        end
    end
end
% warning(oldwarnstat.state,'MATLAB:tex');

end

function wizardprocess
fprintf('\n------------------------------------\n');
fprintf(' Data Dictionary XLS Import Wizard\n');
fprintf('------------------------------------\n');
fprintf('1: Direct import from ''xls'' file(s)\n');
fprintf('2: Open ''txt'' file containing location of target files\n');
fprintf('3: Open directory and specify file importing sequence\n');
fprintf('-- Recent files --\n')
if ispref('DDTool_Config','RecentFiles')
    his=getpref('DDTool_Config','RecentFiles');
    for i=1:numel(his)
        fprintf('%u: %s\n',i+3,his{i});
    end
else
    his={};
end
mxnum=3+numel(his);
fprintf('\n');
sel=input('Selection:  ');
if isnumeric(sel)
    if isempty(sel)
    elseif sel==1
        [filename,pathname]=uigetfile({'*.xls;','Excel Application (*.xls)';'*.*','All Files (*.*)'},'Select files to import','MultiSelect', 'on');
        if isequal(filename,0) || isequal(pathname,0)
            return;
        end
        if iscell(filename)
            multifile_process(strcat(pathname,filename),pathname);
        else
%             hwtbar=waitbar(1,sprintf('Reading "%s"\n',strcat(pathname,filename)));
            read_DD_xls(strcat(pathname,filename));
%             close(hwtbar);
%             fprintf('---->%s\n',strcat(pathname,filename));
        end
    elseif sel==2
        [filename,pathname]=uigetfile({'*.txt;','Text file (*.xls)';'*.*','All Files (*.*)'},'Select file to import');
        if isequal(filename,0) || isequal(pathname,0)
            return;
        end
        xlsfiles=fromstr2xlsfiles(fullfile(pathname,filename),pathname);
%         hwtbar=waitbar(0,'Importing data dictionary');
        for i=1:numel(xlsfiles)
%             waitbar(i/numel(xlsfiles),hwtbar,sprintf('Reading "%s"\n',xlsfiles{i}));
            read_DD_xls(xlsfiles{i});
%             fprintf('---->%s\n',xlsfiles{i});
        end
%         close(hwtbar);
    elseif sel==3
        folderpath=uigetdir;
        if ~isequal(folderpath,0)
            multifile_process(list_xls_under_path(folderpath),folderpath);
        end
    elseif sel>3&&sel<=mxnum
        xlsfiles=fromstr2xlsfiles(his{sel-3});
%         hwtbar=waitbar(0,'Importing data dictionary');
        for i=1:numel(xlsfiles)
%             waitbar(i/numel(xlsfiles),hwtbar,sprintf('Reading "%s"\n',xlsfiles{i}));
            read_DD_xls(xlsfiles{i});
%             fprintf('---->%s\n',xlsfiles{i});
        end
%         close(hwtbar);
    else
        return;
    end
else
    return;
end
end
    




function files=fromstr2xlsfiles(strin,parentpath)
%strin could be xls file, or txt file
%return xls files with its full path
if nargin==1
    parentpath='';
end
files={};
[pathname,filename,ext]=fileparts(strin);
% try to rebuild full path name if only given a name with or without extension
if isempty(ext)
    ext='.xls';
end
if isempty(pathname)
    if isempty(parentpath)
        pathname=pwd;
    else
        pathname=parentpath;
    end
    xlsunderpath=list_xls_under_path(pathname);%if could be found in subdirectory
    canbefound=regexp(xlsunderpath,['\\',filename,ext,'$']);
    for i=1:numel(canbefound)
        if ~isempty(canbefound{i})
            tmpname=xlsunderpath{i};
            [pathname,tmp1,tmp2]=fileparts(tmpname);
            break;
        end
    end
end
fullname=fullfile(pathname,[filename,ext]);%rebuild the fullname

% try to locate the file
if ~exist(fullname,'file')%if the file doesn't exist
    fullname=which([filename,ext]); %search under matlab paths
    if isempty(fullname)
        fprintf('Error: Cannot locate file "%s"\n',strin);
        return;
    end
end
%determine full path name output
if strncmpi(ext,'.xls',4)
    files={fullname};
elseif strcmpi(ext,'.txt')
    fid=fopen(fullname);
    txtlns={};
    while ~feof(fid)
        txtlns=[txtlns;fgetl(fid)];
    end
    for i=1:numel(txtlns)
        files=[files;fromstr2xlsfiles(txtlns{i},pathname)];
    end
end
end



function multifile_process(allxlsfls,folderpath)
allxlsfls=sort_seq(allxlsfls);
if numel(allxlsfls)<1
    return;
end
fprintf('\n-- File list --\n');
for i=1:numel(allxlsfls)
    fprintf('%u: %s\n',i,allxlsfls{i});
end
fprintf('The sequence of xls files listed above has been sorted according to history operation,\n')
idx=askforinput;
    %------Nested function-----------------------
    function idx=askforinput
        inputstr=input('input a new sequence for any necessary change (''c'': Cancel,''h'': Help):  ','s');
        if strcmpi(inputstr,'c')
            idx=0;
            return;
        elseif strcmpi(inputstr,'h')
            fprintf('-- Help --\n');
            tmpstr={'Use the number to represent the file, MATLAB expression is allowed';...
                'For example:'; ...
                'provided there are 10 files listed above'; ...
                '[1 2 3 4 7 5 9] is equal to [1:4 7 5 9], the 6th, 8th and 10th file will be ignored.'};
            fprintf('%s\n',tmpstr{:});
            fprintf('\n');
            idx=askforinput;
        elseif isempty(strfind(inputstr,'['))
            idx=eval(['[',inputstr,']']);
        else
            idx=eval(inputstr);
        end
    end
if ~isempty(idx)
    if isequal(idx,0)
        return;
    end
    xlsfls=allxlsfls(idx);
    update_prefseq(xlsfls);%update sequence history record
    [filename, pathname] = uiputfile('*.txt', 'Save data dictionary group',folderpath);
    if ~(isequal(filename,0) || isequal(pathname,0))
       fid=fopen(fullfile(pathname,filename),'wt');
       fprintf(fid,'%s\n',xlsfls{:});
       fclose(fid);
       save_history(fullfile(pathname,filename));
    end
else
    xlsfls=allxlsfls;
end
% hwtbar=waitbar(0,'Importing data dictionary');
for i=1:numel(xlsfls)
%     waitbar(i/numel(xlsfls),hwtbar,sprintf('Reading "%s"\n',xlsfls{i}));
    read_DD_xls(xlsfls{i});
%     fprintf('---->%s\n',xlsfls{i});
end
% close(hwtbar);
end




function xlsfls_sorted=sort_seq(xlsfls)
%sort the xls files according to history
if numel(xlsfls)<=1
    xlsfls_sorted=xlsfls;
    return;
end
if ispref('DDTool_Config','SequenceHistory')
    xlsfls_sorted={};
    seq=getpref('DDTool_Config','SequenceHistory');
    for i=1:numel(xlsfls)
        [path,flnames{i},ext]=fileparts(xlsfls{i});
    end
    for i=1:numel(seq)
        [TF,LOC]=my_ismember(seq{i},flnames);
        if TF
            xlsfls_sorted=[xlsfls_sorted;xlsfls(LOC)];
        end
    end
    xlsfls_sorted=[xlsfls_sorted;setdiff(xlsfls,xlsfls_sorted)];% unmatched names will be placed at end
else
    xlsfls_sorted=xlsfls;
end
end

function [TF,LOC]=my_ismember(A,S)
LOC=[];
for i=1:numel(S)
    if isequal(A,S{i})
        LOC=[LOC,i];
    end
end
if ~isempty(LOC)
    TF=boolean(1);
else
    TF=boolean(0);
    LOC=0;
end
end

function update_prefseq(xlsfls)
%update the file sequence preference
for i=1:numel(xlsfls)
    [path,flnames{i},ext]=fileparts(xlsfls{i});
end
flnames=flnames';
if ispref('DDTool_Config','SequenceHistory')
    prefseq=getpref('DDTool_Config','SequenceHistory');
    [itsct,ia,ib]=intersect(prefseq,flnames);
    ia=sort(ia);
    ib=sort(ib);
    prefseq(ia)=flnames(ib);%update the same items
    iatmp=[0;ia];ibtmp=[0;ib];
    newprefseq={};
    for i=1:numel(ia)%update the different items
        part=[prefseq(iatmp(i)+1:iatmp(i+1)-1);flnames(ibtmp(i)+1:ibtmp(i+1))];
        newprefseq=[newprefseq;part];
    end
    newprefseq=[newprefseq;flnames(ibtmp(end)+1:end);prefseq(iatmp(end)+1:end)];
    [tmp,iseq]=unique(newprefseq);
    newprefseq=newprefseq(sort(iseq));
    setpref('DDTool_Config','SequenceHistory',newprefseq);
else
    newprefseq=flnames;
    [tmp,iseq]=unique(newprefseq);
    newprefseq=newprefseq(sort(iseq));
    addpref('DDTool_Config','SequenceHistory',newprefseq);
end
end

function c_files=list_xls_under_path(folderpath)
subpaths=path2cell(genpath(folderpath));
c_files={};
for i=1:numel(subpaths)
    fls=dir(subpaths{i});
    tmp={fls.name};
    fls=tmp(~[fls.isdir])';
    exts=regexp(fls,'\.\w+$','match','once');
    xlsfls=fls(strncmpi(exts,'.xls',4));
    if ~isempty(xlsfls)
        xlsfls=strcat(subpaths{i},'\',xlsfls);
    end
    c_files=[c_files;xlsfls];
end
end

function out=path2cell(in);
len=length(in);
sep=strfind(in,';');
cnt=length(sep);
out=cell(1,cnt);
sep=[0,sep];
for i=2:cnt+1
    out{i-1}=in(sep(i-1)+1:sep(i)-1);
end
end

function save_history(filename)
if ispref('DDTool_Config','RecentFiles')
    his=getpref('DDTool_Config','RecentFiles');
    his=[his;filename];
    if numel(his)>10
        his(1:10)=his(end-9:end);
    end
    setpref('DDTool_Config','RecentFiles',his);
else
    addpref('DDTool_Config','RecentFiles',{filename});
end
end




