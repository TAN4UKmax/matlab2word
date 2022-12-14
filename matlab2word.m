classdef matlab2word < handle
    %MATLAB2WORD Transfers calculations to Word report
    %   matlab2word v1.0
    %   Created by TAN4UK
    %   This library can help to transfer your calculations into
    %   Microsoft Word report file.
    %
    %   See detailed description on github page:
    %   https://github.com/TAN4UKmax/matlab2word
    %
    %   MATLAB2WORD Example:
    %   Create new document in Word and type there the following text:
    %
    % String: <string_example> text here
    % Number in formula:
    % a+b+<number_example>=C
    % <number_example>
    % text <number_example> text
    % Figure:
    % <fig_example>
    %
    %   Save your document in *.docx format and close Word application
    %   Then launch the following script in the same folder:
    %
    % %% Calculations
    % test_string = 'Hello';
    % test_num = 0.00000000000000000386-0.000000000013i;
    % % Create test figure
    % x = (-4*pi):(8*pi/1000):(4*pi);
    % f_x_sin_x = sin(x);
    % fig1 = figure(1); % Here you need to stote the figure handler
    % plot(x, f_x_sin_x);
    %
    % %% Write file %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % % Create an instance of matlab2word object and get access to Word file
    % m2w = matlab2word();
    % m2w.SetDecimalSeparator('comma');
    % % Replace variables
    % m2w.Replace('string_example', test_string);
    % m2w.Replace('number_example', test_num);
    % m2w.Replace('fig_example', fig1);
    % % Save Word file
    % m2w.Save();
    %
    %
    %   LICENSE
    %     Copyright (C) 2022 TAN4UK. All Rights Reserved.
    %
    %     Permission is hereby granted, free of charge, to any person
    %     obtaining a copy of this software and associated documentation
    %     files (the "Software"), to deal in the Software without
    %     restriction, including without limitation the rights to use,
    %     copy, modify, merge, publish, distribute, sublicense, and/or sell
    %     copies of the Software, and to permit persons to whom the
    %     Software is furnished to do so, subject to the following
    %     conditions:
    %
    %     The above copyright notice and this permission notice shall be
    %     included in all copies or substantial portions of the Software.
    %
    %     THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    %     EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
    %     OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
    %     NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
    %     HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
    %     WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
    %     FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
    %     OTHER DEALINGS IN THE SOFTWARE.
    
    properties (SetAccess=private)
        % Stores decimal separetor parameter ('dot' by default)
        decimalSeparator
    end
    
    properties (SetAccess=private, Hidden)
        % Stores input file parameters
        inFileSpec
        % Stores Word instance
        word
        % Stores opened document instance
        document
        % Stores cursor position in document
        selection
    end
    
    methods
        function this = matlab2word(filename)
            %MATLAB2WORD Construct an instance of this class
            %   Takes one additional argument as a filename string.
            %   if argument missed, function asks user to select a template
            %   file
            
            % if filename is provided
            if (nargin == 1)
                % Check validity of input file path
                [inFilePath,inFileName,inFileExt] = fileparts(filename);
                if isempty(inFilePath); inFilePath = pwd; end
                if isempty(inFileExt); inFileExt = '.docx'; end
                this.inFileSpec = fullfile(inFilePath, [inFileName, inFileExt]);
            else % if no filename provided
                % Ask user to open file
                [inFileName, inFilePath] = uigetfile( ...
                    {'*.docx', 'Word Documents (*.docx)'}, ...
                    'Select a template document file');
                % If user didn't select a template file
                if (inFileName == 0)
                    % Finish script
                    error('Error! Template file open failed!');
                end
                % Generate full file path
                this.inFileSpec = fullfile(inFilePath, inFileName);
            end
            % Open file
            this.word = actxserver('Word.Application');
            this.word.Visible =1; % shows Word window
            this.document = this.word.Documents.Open(this.inFileSpec);
            this.selection = this.word.Selection;
            % Set default decimal separator to dot
            this.decimalSeparator = 'dot';
        end
        
        function Replace(this, replace_id, replace_data)
            %REPLACE Replaces data in Word
            %   Replace method replaces <replace_id> string in Word
            %   document by replace_data. replace_data can be string,
            %   number or figure instance.
            
            % Add brackets to search  string
            findString = ['<', replace_id, '>'];
            find = this.selection.Find;
            % Search parameters setup
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Text = findString;
            find.Wrap = 1;
            find.MatchCase = true;
            % Search
            find.Execute();
            while find.Found % If found something
                if isfloat(replace_data)
                    % Paste number
                    float_str =  num2str(replace_data);
                    % Replace dot by comma if needed
                    if strcmp(this.decimalSeparator, 'comma')
                        float_str = replace(float_str, '.', ',');
                    end
                    % Fix position of imaginary unit in small or large numbers
                    if (~isreal(replace_data) && (contains(num2str(imag(replace_data)), 'e')))
                        e_index = numel(float_str);
                        while (float_str(e_index) ~= 'e')
                            e_index = e_index - 1;
                        end
                        float_str = [float_str(1:(e_index-1)), 'ie', ...
                            float_str((e_index+1):(end-1))];
                    end
                    if contains(float_str, 'e')
                        % Replace e by power of 10
                        float_str = replace(float_str, 'e', '???10^');
                        this.selection.TypeText(float_str);
                        this.selection.MoveLeft(1, length(float_str), 1);
                        %1=character mode
                        %with this command we mark the previous text%length(text)=amount
                        %1=hold shift
                        this.selection.OMaths.Add(this.selection.Range);
                        this.selection.OMaths.BuildUp();
                    else
                        this.selection.TypeText(float_str);
                    end
                else
                    if isgraphics(replace_data)
                        % Paste figure
                        print(replace_data, '-clipboard', '-dbitmap');
                        this.selection.Paste();
                    else
                        % Paste string
                        this.selection.TypeText(replace_data);
                    end
                end
                % Make search one more time to find all instances
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = findString;
                find.Wrap = 1;
                find.MatchCase = true;
                find.Execute();
            end
        end
        
        function Save(this)
            %SAVE Saves Word file
            %   This method should called in the end of file to save it
            
            % Prepare out file name and extention
            [~, outFileName, fileExt] = fileparts(this.inFileSpec);
            outFileName = append(outFileName, '_out', fileExt);
            % Ask user to confirm output file
            [outFileName, outFilePath] = uiputfile( ...
                {'*.docx', 'Word Documents (*.docx)'}, ...
                'Save As', ...
                outFileName);
            if (outFileName == 0)
                % Finish script
                error('Error! Save file failed!');
            end
            % make output path
            outFileSpec = fullfile(outFilePath, outFileName);
            % Delete old report
            if isfile(outFileSpec)
                delete(outFileSpec);
            end
            % Save new document
            this.document.SaveAs2(outFileSpec);
            
            delete(this);
        end
        
        function SetDecimalSeparator(this, separator)
            %SetDecimalSeparator Sets decimal separator for float numeric
            %data
            %   Dot separator is set by default. You can change separator
            %   to comma by calling this method with argument 'comma'
            if (nargin == 2) && strcmpi(separator, 'comma')
                this.decimalSeparator = 'comma';
            else
                this.decimalSeparator = 'dot';
            end
        end
        
        function delete(this)
            %DELETE Destructor of class
            %   Closes Word and deletes its instance
            
            this.document.Close();  % Close Word file
            this.word.Quit();       % Close Word app
            % Delete Word object (it also deletes all related instances)
            delete(this.word);
            close all;              % Close all opened figures
        end
        
    end
end
