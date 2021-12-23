function varargout=read_DD_xls(fullfilepath)

% ## CONSTANTS ##
IGNORED_SHEETS = {'INPUTS', 'OUTPUTS', 'MAP_DATA'};         % These sheets will be ignored when importing
IGNORED_SHEET_PATTERN = '^#';                   % Sheets with name matches this pattern will be ignored
TITLE_KEYWORDS = {'Name','Min','Max','Unit','Description'}; % "Title row" must contains these keywords
HYPERLINK_TITLE_PATTERN = '^.*Value';    % If title contains the keyword pattern, shall be checked for possible hyperlink

% If no input provided
if nargin<1
    [filename, pathname, filterindex] = uigetfile( ...
        {'*.xls;*.xlsx','Microsoft Excel (*.xls, *.xlsx)'}, ...
        'Select a file');
    if isequal(filename,0) || isequal(pathname,0)
        return;
    end
    fullfilepath=fullfile(pathname,filename);
end
% Start COM server
excel=actxserver('Excel.Application');
bError=false;
set(excel,'Visible',0);
wkbk=excel.Workbooks.Open(fullfilepath);
fprintf('## Reading from "%s"...\n',fullfilepath);
shts=wkbk.Sheets;
% context_info is used to pass context information
context_info.file=fullfilepath; % update context_info variable
context_info.project = dd_getproject;

for isht=1:shts.Count % Iterates worksheets
    sht=shts.Item(isht);
    context_info.sheet=sht.Name; %update context_info variable
    % Skip unused sheet
    if sht.UsedRange.Count<=1 %if sheet is blank
        continue;
    elseif ismember(sht.Name,IGNORED_SHEETS) || ~isempty(regexp(sht.Name, IGNORED_SHEET_PATTERN))
        continue;
    else
        raw_contents=sht.UsedRange.Value; %raw contents
        raw_contents=[cell(sht.UsedRange.Item(1).Row-1,size(raw_contents,2));raw_contents];
        % Note: Title row could be preceded with information rows at will
        row_indexer=1;
        % locate the title row first
        while row_indexer<size(raw_contents,1)
            titlerow=cellfun(@num2str,raw_contents(row_indexer,:),'UniformOutput', 0);
            titlerow=strtrim(titlerow);
            if isempty(setdiff(TITLE_KEYWORDS,titlerow)) %All keywords matched, title row located successfully
                break;
            else
                row_indexer=row_indexer+1;
            end
        end
        if row_indexer>size(raw_contents,1)
            err_info = 'title_locate_fail';
            bError=true;
            print_error_info(err_info,context_info);
        end
        % reformat title row so as to be used as valid struct field name
        titlerow_raw = titlerow; % backup for later use
        for ititle=1:numel(titlerow)
            if ~isempty(titlerow{ititle})
                titlerow{ititle}=genvarname(titlerow{ititle});
            end
        end
        % begin row traversal
        for ir=row_indexer+1:size(raw_contents,1)
            context_info.row=ir; %update context_info variable
            for ititle=1:numel(titlerow)
                if ~isempty(titlerow_raw{ititle})
                    varobject.(titlerow{ititle})=raw_contents{ir,ititle};
                end
            end
            if isempty(varobject.Name)||~isstr(varobject.Name)
                continue;
            end
            varobject.Name=strtrim(varobject.Name);
            context_info.var=varobject.Name;%update context_info variable
            % Further evaluate value range
            for i=1:numel(titlerow)
                if ~isempty(regexp(titlerow_raw{i},HYPERLINK_TITLE_PATTERN))
                    valrng=sht.Range([dec2base27(i), int2str(ir)]);
                    [varobject.(titlerow{i}), err_info]=evaluate_value(valrng, context_info);
                    if ~isempty(err_info)
                        bError=true;
                        print_error_info(err_info,context_info);
                    end
                end
            end
            %Type conversion for properties
            % !!! Temporarily convert all non-empty fields to string, to be improved
            flds=fieldnames(varobject);
            for i=1:numel(flds)
                if ~isstr(varobject.(flds{i}))&&numel(varobject.(flds{i}))==1&&isnan(varobject.(flds{i}))
                    % convert NaN
                    varobject.(flds{i})='';
                end
            end
            % Call dispatcher to actually generate data object
            varobject_dispatcher(varobject, context_info, false);
        end
    end
end
excel.Quit;
if nargout>0
    varargout={bError};
end
if nargin<1
    if bError
        errordlg('Error occured during the process, refer to screen for detail.');
    else
        msgbox('Import successfully finished.');
    end
end
end % main function end



% ##############################
function [val, err_info] = evaluate_value(valrng, context_info)
% Further analyze the target field:
% Try to evaluate in sequence: 
%   number/hyperlink/evaluate/invalid
%                    evaluate -> data object/not data object
sht = valrng.Worksheet;
val = valrng.Value;
excel = valrng.Application;
err_info = '';
if isstr(val) %if 'val' is string
    if valrng.Hyperlinks.Count>0 %if hyperlink exists
        %Go to hyperlink to read
        base_addr=valrng.Hyperlinks.Item(1).SubAddress;
        base_rng=excel.Range(base_addr);
        if base_rng.Column~=1 ...  % Target must be on the first column
           ||base_rng.Count>1 ...  % Cannot be merged cells
           ||~strcmpi(base_rng.Value,context_info.var) % Name must match
            err_info = 'hyperlink_invalid';
            val=[];return;
        else
            mapsht=base_rng.Worksheet;
            startrow=base_rng.Row+1; %starting row of data block
            data_lt=['A',int2str(startrow)]; % Left top of data block
            % Search for column bound
            colcnt=0;
            rngdata=mapsht.Range(data_lt).Value;
            while (isnumeric(rngdata)&&~isnan(rngdata)) ... % numeric and not NaN
                    ||(isstr(rngdata)&&~isempty(evaluate_val_str(rngdata))) %string and could be converted to number
                colcnt=colcnt+1;
                rngdata=mapsht.Range([dec2base27(1+colcnt),int2str(startrow)]).Value; %pointer move right
            end
            % Search for row bound
            rowcnt=0;
            rngdata=mapsht.Range(data_lt).Value;
            while (isnumeric(rngdata)&&~isnan(rngdata)) ... % numeric and not NaN
                    ||(isstr(rngdata)&&~isempty(evaluate_val_str(rngdata))) %string and could be converted to number
                rowcnt=rowcnt+1;
                rngdata=mapsht.Range(['A',int2str(startrow+rowcnt)]).Value; %pointer move right
            end
            %extract data block
            datablk=mapsht.Range([data_lt,':',dec2base27(colcnt),int2str(startrow+rowcnt-1)]).Value;
            for kk=1:numel(datablk)
                if isstr(datablk{kk})
                    [datablk{kk}, err_info]=evaluate_val_str(datablk{kk});
                    if ~isempty(err_info)
                        val=[];return;
                    end
                end
            end
            try
                val=cell2mat(datablk);
            catch
                err_info = 'data_block_invalid';
                val=[];return;
            end
        end
    else % not a hyperlink, should be evaluated
        [val, err_info] = evaluate_val_str(val);
        if ~isempty(err_info)
            return;
        end
    end
else % if not a string
    % nothing need to do
end
end


% ##############################
function [val, err_info] = evaluate_val_str(val_str)
err_info = '';
if ~isstr(val_str)
    val = val_str;
    return;
end
try %try to evaluate the expression
    rawval=val_str;
    val=evalin('base',val_str);
    if isa(val,'Simulink.Parameter') || isa(val,'Simulink.Signal') %defined as data object
        val=double(val.Value); % must convert it to double because other values from Excel is by default double
    elseif isenum(val)
        val = val_str;
    else
    end
catch
    try
        valstr=regexprep(valstr,'(\<\D\w+)','$1.Value');
        val=evalin('base',valstr);
    catch
        val = val_str;
        err_info = 'evaluation_fail';
    end
end
end

function tf = isenum(val)
f = metaclass(val);
tf = f.Enumeration;
end

% ##############################
function print_error_info(err_info, context_info)
switch err_info
    case 'title_locate_fail'
        fprintf('## Error: Failed to locate title row from %s, invalid format\n',context_info.sheet);
    case 'hyperlink_invalid'
        fprintf('## Error: %s has an invalid hyperlink destination (Sheet: %s, Line: %u)\n',context_info.var,context_info.sheet,context_info.row);
    case 'data_block_invalid'
        fprintf('## Error: %s followed by invalid data block (Sheet: %s, Line: %u)\n',context_info.var,context_info.sheet,context_info.row);
    case 'evaluation_fail'
        fprintf('## Error: %s cannot be evaluated (Sheet: %s, Line: %u)\n',context_info.var,context_info.sheet,context_info.row);
    otherwise
end
end


%% Supportive functions
%------------------------------------------------------------------------------
function s = dec2base27(d)

%   DEC2BASE27(D) returns the representation of D as a string in base 27,
%   expressed as 'A'..'Z', 'AA','AB'...'AZ', and so on. Note, there is no zero
%   digit, so strictly we have hybrid base26, base27 number system.  D must be a
%   negative integer bigger than 0 and smaller than 2^52.
%
%   Examples
%       dec2base(1) returns 'A'
%       dec2base(26) returns 'Z'
%       dec2base(27) returns 'AA'
%-----------------------------------------------------------------------------

d = d(:);
if d ~= floor(d) || any(d(:) < 0) || any(d(:) > 1/eps)
    error('MATLAB:xlswrite:Dec2BaseInput',...
        'D must be an integer, 0 <= D <= 2^52.');
end
[num_digits begin] = calculate_range(d);
s = index_to_string(d, begin, num_digits);
end

function string = index_to_string(index, first_in_range, digits)

letters = 'A':'Z';
working_index = index - first_in_range;
outputs = cell(1,digits);
[outputs{1:digits}] = ind2sub(repmat(26,1,digits), working_index);
string = fliplr(letters([outputs{:}]));
end
%----------------------------------------------------------------------
function [digits first_in_range] = calculate_range(num_to_convert)

digits = 1;
first_in_range = 0;
current_sum = 26;
while num_to_convert > current_sum
    digits = digits + 1;
    first_in_range = current_sum;
    current_sum = first_in_range + 26.^digits;
end
end


function prjname = dd_getproject
if isempty(bdroot)
    prjname = '';
    return;
end
blk = find_system(bdroot, 'FollowLinks', 'on', 'LookUnderMasks', 'on', 'MaskType', 'DDTools_Project');
if isempty(blk)
    prjname = ''; return;
elseif numel(blk)==1
    prjname = get_param(blk, 'project_name');
else %numel(blk)>1
    prjname = unique(get_param(blk, 'project_name'));
    if numel(prjname)>1
        error('Multiple <DDTools_Project> blocks with different project names found inside current model');
    end
end
if iscell(prjname)
    prjname = prjname{1};
end

end