namespace MatrixArmanshin
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            MatrixA = new DataGridView();
            MatrixSize = new ComboBox();
            label1 = new Label();
            ResultButton = new Button();
            ClearButton = new Button();
            ExitButton = new Button();
            PDFButton = new Button();
            ExcelButton = new Button();
            WordButton = new Button();
            resultLabel = new Label();
            ((System.ComponentModel.ISupportInitialize)MatrixA).BeginInit();
            SuspendLayout();
            // 
            // MatrixA
            // 
            MatrixA.AllowUserToAddRows = false;
            MatrixA.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            MatrixA.Location = new Point(64, 43);
            MatrixA.Name = "MatrixA";
            MatrixA.RowTemplate.Height = 25;
            MatrixA.Size = new Size(319, 165);
            MatrixA.TabIndex = 0;
            // 
            // MatrixSize
            // 
            MatrixSize.FormattingEnabled = true;
            MatrixSize.Location = new Point(389, 43);
            MatrixSize.Name = "MatrixSize";
            MatrixSize.Size = new Size(52, 23);
            MatrixSize.TabIndex = 2;
            MatrixSize.SelectedIndexChanged += MatrixSize_SelectedIndexChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 14.25F, FontStyle.Bold, GraphicsUnit.Point);
            label1.Location = new Point(174, 15);
            label1.Name = "label1";
            label1.Size = new Size(95, 25);
            label1.TabIndex = 3;
            label1.Text = "Матрица";
            // 
            // ResultButton
            // 
            ResultButton.Location = new Point(227, 225);
            ResultButton.Name = "ResultButton";
            ResultButton.Size = new Size(156, 50);
            ResultButton.TabIndex = 5;
            ResultButton.Text = "Рассчитать";
            ResultButton.UseVisualStyleBackColor = true;
            ResultButton.Click += ResultButton_Click;
            // 
            // ClearButton
            // 
            ClearButton.Location = new Point(227, 281);
            ClearButton.Name = "ClearButton";
            ClearButton.Size = new Size(156, 50);
            ClearButton.TabIndex = 6;
            ClearButton.Text = "Очистить";
            ClearButton.UseVisualStyleBackColor = true;
            ClearButton.Click += ClearButton_Click;
            // 
            // ExitButton
            // 
            ExitButton.Location = new Point(227, 337);
            ExitButton.Name = "ExitButton";
            ExitButton.Size = new Size(156, 50);
            ExitButton.TabIndex = 7;
            ExitButton.Text = "Выход";
            ExitButton.UseVisualStyleBackColor = true;
            ExitButton.Click += ExitButton_Click;
            // 
            // PDFButton
            // 
            PDFButton.Location = new Point(65, 337);
            PDFButton.Name = "PDFButton";
            PDFButton.Size = new Size(156, 50);
            PDFButton.TabIndex = 10;
            PDFButton.Text = "Вывести в PDF";
            PDFButton.UseVisualStyleBackColor = true;
            PDFButton.Click += PDFButton_Click;
            // 
            // ExcelButton
            // 
            ExcelButton.Location = new Point(65, 281);
            ExcelButton.Name = "ExcelButton";
            ExcelButton.Size = new Size(156, 50);
            ExcelButton.TabIndex = 9;
            ExcelButton.Text = "Вывести в Excel";
            ExcelButton.UseVisualStyleBackColor = true;
            ExcelButton.Click += ExcelButton_Click;
            // 
            // WordButton
            // 
            WordButton.Location = new Point(65, 225);
            WordButton.Name = "WordButton";
            WordButton.Size = new Size(156, 50);
            WordButton.TabIndex = 8;
            WordButton.Text = "Вывести в Word";
            WordButton.UseVisualStyleBackColor = true;
            WordButton.Click += WordButton_Click;
            // 
            // resultLabel
            // 
            resultLabel.AutoSize = true;
            resultLabel.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            resultLabel.Location = new Point(389, 99);
            resultLabel.Name = "resultLabel";
            resultLabel.Size = new Size(0, 15);
            resultLabel.TabIndex = 11;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(818, 419);
            Controls.Add(resultLabel);
            Controls.Add(PDFButton);
            Controls.Add(ExcelButton);
            Controls.Add(WordButton);
            Controls.Add(ExitButton);
            Controls.Add(ClearButton);
            Controls.Add(ResultButton);
            Controls.Add(label1);
            Controls.Add(MatrixSize);
            Controls.Add(MatrixA);
            Name = "Form1";
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)MatrixA).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView MatrixA;
        private ComboBox MatrixSize;
        private Label label1;
        private Button ResultButton;
        private Button ClearButton;
        private Button ExitButton;
        private Button PDFButton;
        private Button ExcelButton;
        private Button WordButton;
        private Label resultLabel;
    }
}