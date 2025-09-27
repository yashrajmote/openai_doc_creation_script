#!/usr/bin/env python3
"""
Drag and Drop Desktop Application for Excel to DOCX Generator
A simple drag-and-drop interface using tkinter with file drop functionality.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import tkinterdnd2 as tkdnd
import os
import threading
import openai
from dotenv import load_dotenv
from excel_to_docx_generator import ExcelToDocxGenerator

# Load environment variables
load_dotenv()


class DragDropApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to DOCX Generator - Drag & Drop")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # Variables
        self.output_directory = tk.StringVar()
        self.resume_output_directory = tk.StringVar()
        self.processing = False
        
        # Hardcoded template paths - Update these to your actual template file paths
        self.java_template_path = "/Users/yash/Desktop/openai_doc_creation_script/templates/java_resume_template.dotx"
        self.csharp_template_path = "/Users/yash/Desktop/openai_doc_creation_script/templates/csharp_resume_template.dotx"
        
        # OpenAI API key from environment variable
        self.openai_api_key = os.getenv('OPENAI_API_KEY')
        if not self.openai_api_key:
            messagebox.showerror("Error", "OpenAI API key not found. Please set OPENAI_API_KEY environment variable or create a .env file.")
            self.root.quit()
            return
        
        # Resume templates
        self.java_resume = """SKILLS -	Languages: Java, C#, Python, C++, JavaScript, TypeScript -	Frameworks and Libraries: Spring Boot, Angular, React, Next.js, Tailwind CSS, Material-UI -	Databases & Tools: MySQL, PostgreSQL, MongoDB, Firebase -	DevOps and Cloud: Docker, GitHub Actions, AWS (S3, Lambda), Git, CI/CD -	Other Tools: Kafka, Jenkins, Jira, Sanity CMS, Jasmine (Unit Testing) WORK EXPERIENCE Abhitech Energycon Limited, Toledo, OH							  	 May 2024 – Dec 2024 Full Stack Software Engineer Intern •	Developed a Gain/Loss Dashboard for coal power plants using Angular and TailwindCSS, helping users identify and act on data patterns contributing to monthly operational losses. •	Built an ETL pipeline with Apache Kafka to extract data from SAP into MySQL, improved query speed by 50% with indexing. •	Implemented Docker for containerization and integrated with GitHub Actions for CI/CD pipeline. Leveraged AWS S3 to store large volumes of SAP data and AWS Lambda to automate data processing, reducing manual intervention. Crown Equipment, New Bremen, OH	           August 2023 – Dec 2023 Full Stack Software Engineer intern •	Developed an internal invoicing application using Angular, Typescript and designed RESTful APIs in Spring Boot, Java to reduce invoice processing time by 90%. Built a reusable Radio Button List component to improve form input UX across the invoice app. •	Utilized SQL queries and stored procedures to process large datasets within the invoicing platform, resulting in 50% faster financial reporting. •	Implemented Unit Tests using Jasmine to validate Angular components, identified critical edge-cases and performed debugging, achieving a 95% test coverage. •	Collaborated with cross-functional teams in the Software Development Life Cycle in an Agile/Scrum environment, performing Code Review and QA through Jira-based workflows, and maintained technical documentation. Technochrafts, remote	  January 2023 – July 2023 Software Engineer intern •	Developed a secure cross-platform registration and login system using Java and Spring Boot to create RESTful APIs. Integrated OAuth 2.0 JWT authentication using Spring Security to ensure session management. •	Designed and implemented a responsive, user-friendly webpage using HTML, CSS, and JavaScript, ensuring cross-browser compatibility. Configured and managed the web server with NGINX to enhance performance, load balancing, and server-side caching. •	Implemented a CI/CD pipeline using GitLab CI for automated testing, integration, and deployment. Deployed containerized applications to AWS via AWS Elastic Beanstalk for simplified deployment and auto scaling of app. PROJECTS ML Intern – Anonymous Insurance Company •	Built an automated MLOps using Docker and MLFlow to retrain and evaluate XGBoost models on insurance datasets. The pipeline optimized model performance, reduced tuning time by 25%, and seamlessly handled updates for datasets exceeding 10 million rows. Headstarter Fellowship – Pantry Tracker App | Next.js, React, Firebase •	Developed a web-based inventory system using Next.js and Firebase, implemented real-time updates and item categorization using Firestore listeners. Headstarter Fellowship – AI Customer Support •	Built a real-time AI-powered chat assistant using OpenAI API and Next.js, backed by AWS Lambda and WebSockets to handle 10K+ concurrent requests with <200ms latency. EDUCATION University of Toledo, Toledo, OH Bachelor of Science Degree Recipient | GPA – 3.3 | Major: Computer Science Engineering Honors and Awards: Dean's List (2020 – 2022), UToledo Rockets Scholarship, Engineering Scholarship"""

        self.csharp_resume = """YASHRAJ MOTE LinkedIn | ymote@rockets.utoledo.edu |  GitHub SKILLS -	Languages: C#, Java, Python, C++, JavaScript, TypeScript -	Frameworks and Libraries: ASP.NET Core, Entity Framework Core, Spring Boot -	Frontend: Blazor, Angular, React, Next.js, Tailwind CSS, HTML5, CSS3 -	Databases & Tools: Microsoft SQL Server, MySQL, PostgreSQL, MongoDB, Firebase -	DevOps and Cloud: Docker, Kubernetes, GitHub Actions, GitLab CI, AWS (S3, Lambda), Azure, Git, CI/CD -	Other Tools: Kafka, Jenkins, Jira, Sanity CMS, Jasmine (Unit Testing) WORK EXPERIENCE Abhitech Energycon Limited, Toledo, OH							  	 May 2024 – Dec 2024 Full Stack Software Engineer Intern •	Developed a Gain/Loss Dashboard for coal power plants using Blazor, ASP.NET Core Web API and Entity Framework Core, helping users identify and act on data patterns contributing to monthly operational losses. •	Built an ETL pipeline with Apache Kafka to extract data from SAP into Microsoft SQL Server, and improved query speed by 50% with indexing. •	Implemented Docker for containerization and integrated with GitHub Actions for CI/CD pipeline. Leveraged AWS S3 to store large volumes of SAP data and AWS Lambda to automate data processing, reducing manual intervention. Crown Equipment, New Bremen, OH	           August 2023 – Dec 2023 Full Stack Software Engineer intern •	Developed an internal invoicing application using Angular, Typescript and designed RESTful APIs in Spring Boot, Java to reduce invoice processing time by 90%. Built a reusable Radio Button List component to improve form input UX. •	Utilized SQL queries and stored procedures to process large datasets within the invoicing platform, resulting in 50% faster financial reporting. •	Implemented Unit Tests using Jasmine to validate Angular components, identified critical edge-cases and performed debugging, achieving a 95% test coverage. •	Collaborated with cross-functional teams in the Software Development Life Cycle in an Agile/Scrum environment, performing Code Review and QA through Jira-based workflows, and maintained technical documentation. Technochrafts, remote	  January 2023 – July 2023 Software Engineer intern •	Developed a secure login system using ASP.NET Core and Entity Framework Core, implementing OAuth 2.0 and JWT for authentication and session management. •	Designed and implemented RESTful APIs in the backend service. Deployed and configured NGINX on Azure App Service to enable load balancing, implement server-side caching and optimize performance by 45%. •	Designed a CI/CD pipeline with GitLab CI and Docker, deploying to Azure Kubernetes Service to automate testing and cut deployment time by 40%. PROJECTS ML Intern – Anonymous Insurance Company •	Designed an MLOps pipeline with Docker and MLflow to automate model retraining and evaluation for XGBoost on insurance datasets. The pipeline optimized model performance, reduced tuning time by 25%, and seamlessly handled updates for datasets exceeding 10 million rows. Headstarter Fellowship – Pantry Tracker App | Next.js, React, Firebase •	Developed a web-based inventory system using Next.js and Firebase, implemented real-time updates and item categorization using Firestore listeners. Headstarter Fellowship – AI Customer Support •	Built a real-time AI-powered chat assistant using OpenAI API and Next.js, backed by AWS Lambda and WebSockets to handle 10K+ concurrent requests with <200ms latency. EDUCATION University of Toledo, Toledo, OH Bachelor of Science Degree Recipient | GPA – 3.3 | Major: Computer Science Engineering Honors and Awards: Dean's List (2020 – 2022), UToledo Rockets Scholarship, Engineering Scholarship"""
        
        # Template keywords for detection
        self.java_keywords = {
            'java', 'spring', 'spring boot', 'maven', 'gradle', 'hibernate', 
            'junit', 'tomcat', 'jpa', 'jdbc', 'servlet', 'jsp', 'struts'
        }
        
        self.csharp_keywords = {
            'c#', 'csharp', '.net', 'asp.net', 'entity framework', 'mvc', 
            'blazor', 'azure', 'sql server', 'visual studio', 'nuget', 'wcf'
        }
        
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="📄 Excel to DOCX Generator", 
                               font=("Arial", 18, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Template Status Section
        template_frame = ttk.LabelFrame(main_frame, text="Resume Templates", padding="10")
        template_frame.pack(fill=tk.X, pady=10)
        
        # Template status labels
        java_status = "✅ Found" if os.path.exists(self.java_template_path) else "❌ Not Found"
        csharp_status = "✅ Found" if os.path.exists(self.csharp_template_path) else "❌ Not Found"
        
        ttk.Label(template_frame, text=f"Java Template: {java_status}").pack(anchor=tk.W)
        ttk.Label(template_frame, text=f"C# Template: {csharp_status}").pack(anchor=tk.W)
        
        if not os.path.exists(self.java_template_path) or not os.path.exists(self.csharp_template_path):
            ttk.Label(template_frame, text="⚠️ Please ensure template files exist at the specified paths", 
                     foreground="red").pack(anchor=tk.W, pady=5)
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                               text="Drag and drop your Excel file here\nor click to browse",
                               font=("Arial", 12),
                               foreground="gray")
        instructions.pack(pady=10)
        
        # Process button - MOVED ABOVE DROP ZONE
        self.process_button = ttk.Button(main_frame, text="🚀 Generate DOCX Files", 
                                       command=self.process_files, 
                                       state="disabled",
                                       style="Accent.TButton")
        self.process_button.pack(pady=15)
        
        # Output directory
        dir_frame = ttk.Frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(dir_frame, text="DOCX Output Directory:").pack(side=tk.LEFT)
        ttk.Entry(dir_frame, textvariable=self.output_directory, width=30).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Button(dir_frame, text="Browse", command=self.browse_output_directory).pack(side=tk.LEFT)
        
        # Resume output directory
        resume_dir_frame = ttk.Frame(main_frame)
        resume_dir_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(resume_dir_frame, text="Resume Output Directory:").pack(side=tk.LEFT)
        ttk.Entry(resume_dir_frame, textvariable=self.resume_output_directory, width=30).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Button(resume_dir_frame, text="Browse", command=self.browse_resume_output_directory).pack(side=tk.LEFT)
        
        # Drop zone
        self.drop_zone = tk.Frame(main_frame, 
                                 bg="lightgray", 
                                 relief="ridge", 
                                 bd=2,
                                 height=120)
        self.drop_zone.pack(fill=tk.BOTH, expand=True, pady=10)
        self.drop_zone.pack_propagate(False)
        
        # Drop zone label
        self.drop_label = ttk.Label(self.drop_zone, 
                                   text="📁 Drop Excel file here",
                                   font=("Arial", 14),
                                   background="lightgray")
        self.drop_label.pack(expand=True)
        
        # File browser button inside drop zone
        browse_button = ttk.Button(self.drop_zone, 
                                  text="📂 Or click to browse",
                                  command=self.browse_excel_file,
                                  style="Accent.TButton")
        browse_button.pack(pady=10)
        
        # Configure drag and drop
        self.drop_zone.drop_target_register(tkdnd.DND_FILES)
        self.drop_zone.dnd_bind('<<Drop>>', self.on_drop)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Ready - Drop an Excel file to begin")
        self.status_label.pack()
        
        # Results area
        self.results_text = scrolledtext.ScrolledText(main_frame, height=8, width=60)
        self.results_text.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Set default output directories
        self.output_directory.set("/Users/yash/Desktop/Desktop - YASH's MacBook Air/APPLICATIONS")
        self.resume_output_directory.set("/Users/yash/Desktop/Desktop - YASH's MacBook Air/APPLICATIONS")
        
        # Store current file
        self.current_file = None
        
    def on_drop(self, event):
        """Handle file drop event."""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            if file_path.lower().endswith(('.xlsx', '.xls')):
                self.current_file = file_path
                self.drop_label.config(text=f"✅ {os.path.basename(file_path)}")
                self.process_button.config(state="normal")
                self.status_label.config(text="File ready - Click 'Generate DOCX Files' to process")
            else:
                messagebox.showerror("Error", "Please drop an Excel file (.xlsx or .xls)")
                
    def browse_excel_file(self):
        """Browse for Excel file."""
        from tkinter import filedialog
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.current_file = filename
            self.drop_label.config(text=f"✅ {os.path.basename(filename)}")
            self.process_button.config(state="normal")
            self.status_label.config(text="File ready - Click 'Generate DOCX Files' to process")
            
    def browse_output_directory(self):
        """Browse for output directory."""
        from tkinter import filedialog
        directory = filedialog.askdirectory(title="Select DOCX Output Directory")
        if directory:
            self.output_directory.set(directory)
            
    def browse_resume_output_directory(self):
        """Browse for resume output directory."""
        from tkinter import filedialog
        directory = filedialog.askdirectory(title="Select Resume Output Directory")
        if directory:
            self.resume_output_directory.set(directory)
            
    
    def detect_template_type(self, job_description, position):
        """Detect which template to use based on job description and position."""
        text = f"{job_description} {position}".lower()
        
        java_score = sum(1 for keyword in self.java_keywords if keyword in text)
        csharp_score = sum(1 for keyword in self.csharp_keywords if keyword in text)
        
        # If scores are tied or both zero, use position title as tiebreaker
        if java_score == csharp_score:
            if any(keyword in text for keyword in ['java', 'spring']):
                return 'java'
            elif any(keyword in text for keyword in ['c#', '.net', 'csharp']):
                return 'csharp'
            else:
                return 'java'  # Default to Java
        
        return 'java' if java_score > csharp_score else 'csharp'
    
    def load_template(self, template_path):
        """Load a .dotx template file."""
        try:
            # Load the template file (works for both .dotx and .docx)
            doc = Document(template_path)
            return doc
        except Exception as e:
            raise Exception(f"Error loading template {template_path}: {str(e)}")
    
    def populate_template(self, doc, resume_content, company, position):
        """Populate template with AI-generated content while preserving template structure."""
        try:
            # Create a new document based on the template
            new_doc = Document()
            
            # Copy the template's styles and formatting
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    new_para = new_doc.add_paragraph()
                    new_para.style = paragraph.style
                    new_para.text = paragraph.text
            
            # Add company and position info at the top
            new_doc.add_heading(f'Tailored Resume for {company}', level=1)
            new_doc.add_paragraph(f'Position: {position}')
            new_doc.add_paragraph('')  # Empty line
            
            # Add AI-generated content with proper formatting
            sections = resume_content.split('\n\n')
            for section in sections:
                if section.strip():
                    # Check if it's a heading (all caps or starts with common section names)
                    if (section.strip().isupper() and len(section.strip()) < 50) or \
                       any(section.strip().upper().startswith(heading) for heading in 
                           ['SKILLS', 'WORK EXPERIENCE', 'EDUCATION', 'PROJECTS']):
                        new_doc.add_heading(section.strip(), level=2)
                    else:
                        new_doc.add_paragraph(section.strip())
            
            return new_doc
            
        except Exception as e:
            raise Exception(f"Error populating template: {str(e)}")
            
    def process_files(self):
        """Process the Excel file and generate DOCX files."""
        if not self.current_file:
            messagebox.showerror("Error", "Please drop an Excel file first.")
            return
            
        if not self.output_directory.get():
            messagebox.showerror("Error", "Please select an output directory.")
            return
            
        if not os.path.exists(self.current_file):
            messagebox.showerror("Error", "Excel file does not exist.")
            return
            
        # Start processing in a separate thread
        self.processing = True
        self.process_button.config(state="disabled")
        self.status_label.config(text="Processing...")
        self.results_text.delete(1.0, tk.END)
        
        # Run processing in separate thread
        thread = threading.Thread(target=self.run_processing)
        thread.daemon = True
        thread.start()
        
    def run_processing(self):
        """Run the actual processing in a separate thread."""
        try:
            # Create generator and process file
            generator = ExcelToDocxGenerator(
                self.current_file, 
                self.output_directory.get()
            )
            
            results = generator.process_excel_file()
            
            if results:
                # Generate resumes using OpenAI
                self.root.after(0, self.update_results, results)
                self.root.after(0, self.generate_resumes, results)
            else:
                self.root.after(0, self.show_error, "Processing failed. Please check the error messages.")
                
        except Exception as e:
            self.root.after(0, self.show_error, f"Error: {str(e)}")
        finally:
            self.root.after(0, self.processing_complete)
            
    def generate_resumes(self, results):
        """Generate tailored resumes using OpenAI API and templates."""
        try:
            # Set up OpenAI client
            client = openai.OpenAI(api_key=self.openai_api_key)
            
            # Read the Excel file to get job descriptions
            import pandas as pd
            df = pd.read_excel(self.current_file, header=12)
            df.columns = df.columns.str.strip()
            
            # Find the job description column (similar to how we find Company and Position)
            job_desc_col = None
            for col in df.columns:
                col_lower = col.lower().strip()
                if 'description' in col_lower or 'job description' in col_lower or 'job_desc' in col_lower:
                    job_desc_col = col
                    break
            
            if not job_desc_col:
                self.results_text.insert(tk.END, "\n⚠️ No job description column found. Skipping resume generation.\n")
                return
            
            # Check if templates exist
            use_templates = os.path.exists(self.java_template_path) and os.path.exists(self.csharp_template_path)
            
            if use_templates:
                self.results_text.insert(tk.END, "\n🎨 Using template-based resume generation...\n")
                # Generate individual resume files using templates
                self.generate_individual_resumes_with_templates(client, df, job_desc_col)
                # Also generate master document with all responses
                self.results_text.insert(tk.END, "\n📄 Generating master document...\n")
                self.generate_combined_resume(client, df, job_desc_col)
            else:
                self.results_text.insert(tk.END, "\n📝 Using text-based resume generation (templates not found)...\n")
                self.generate_combined_resume(client, df, job_desc_col)
            
        except Exception as e:
            self.results_text.insert(tk.END, f"\n❌ Error in resume generation: {str(e)}\n")
    
    def generate_individual_resumes_with_templates(self, client, df, job_desc_col):
        """Generate individual resume files using templates."""
        valid_entries = 0
        total_entries = len(df)
        
        self.results_text.insert(tk.END, f"\n🔍 Found {total_entries} total rows. Generating individual resumes...\n")
        self.results_text.see(tk.END)
        self.root.update()
        
        for index, row in df.iterrows():
            company = row.get('Company', '')
            position = row.get('Position', '')
            job_description = row.get(job_desc_col, '')
            
            # Skip if any required field is empty
            if pd.isna(company) or pd.isna(position) or pd.isna(job_description) or not str(job_description).strip():
                continue
            
            valid_entries += 1
            self.results_text.insert(tk.END, f"\n🤖 [{valid_entries}] Generating resume for {company}...\n")
            self.results_text.see(tk.END)
            self.root.update()
            
            try:
                # Detect which template to use
                template_type = self.detect_template_type(str(job_description), str(position))
                template_path = self.java_template_path if template_type == 'java' else self.csharp_template_path
                
                self.results_text.insert(tk.END, f"📋 Using {template_type.upper()} template\n")
                self.results_text.see(tk.END)
                self.root.update()
                
                # Generate resume content using OpenAI
                resume_content = self.generate_single_resume(client, company, position, str(job_description), template_type)
                
                # Load and populate template
                doc = self.load_template(template_path)
                populated_doc = self.populate_template(doc, resume_content, company, position)
                
                # Save individual resume file
                import re
                safe_company = re.sub(r'[<>:"/\\|?*]', '_', str(company))[:30]
                safe_position = re.sub(r'[<>:"/\\|?*]', '_', str(position))[:30]
                resume_filename = f"{valid_entries:03d}_{safe_company}_{safe_position}_Resume.docx"
                resume_filepath = os.path.join(self.resume_output_directory.get(), resume_filename)
                
                populated_doc.save(resume_filepath)
                
                self.results_text.insert(tk.END, f"✅ Created: {resume_filename}\n")
                self.results_text.see(tk.END)
                self.root.update()
                
            except Exception as e:
                self.results_text.insert(tk.END, f"❌ Error generating resume for {company}: {str(e)}\n")
                continue
        
        self.results_text.insert(tk.END, f"\n✅ Generated {valid_entries} individual tailored resumes!\n")
    
    def generate_combined_resume(self, client, df, job_desc_col):
        """Generate a combined resume document (original behavior)."""
        from docx import Document
        from docx.shared import Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        
        doc = Document()
        doc.add_heading('AI-Generated Tailored Resumes', 0)
        
        # Process each valid entry
        valid_entries = 0
        total_entries = len(df)
        
        self.results_text.insert(tk.END, f"\n🔍 Found {total_entries} total rows. Processing job descriptions...\n")
        self.results_text.see(tk.END)
        self.root.update()
        
        for index, row in df.iterrows():
            company = row.get('Company', '')
            position = row.get('Position', '')
            job_description = row.get(job_desc_col, '')
            
            # Skip if any required field is empty
            if pd.isna(company) or pd.isna(position) or pd.isna(job_description) or not str(job_description).strip():
                continue
            
            valid_entries += 1
            self.results_text.insert(tk.END, f"\n🤖 [{valid_entries}] Generating resume for {company}...\n")
            self.results_text.see(tk.END)
            self.root.update()
            
            try:
                # Generate resume using OpenAI
                resume_content = self.generate_single_resume(client, company, position, str(job_description), 'java')
                
                # Add to document
                doc.add_heading(f'{company} - {position}', level=1)
                doc.add_paragraph(resume_content)
                doc.add_page_break()
                
                self.results_text.insert(tk.END, f"✅ Completed {company}\n")
                self.results_text.see(tk.END)
                self.root.update()
                
            except Exception as e:
                self.results_text.insert(tk.END, f"❌ Error generating resume for {company}: {str(e)}\n")
                continue
        
        # Save the resume document
        resume_file_path = os.path.join(self.resume_output_directory.get(), "AI_Generated_Resumes.docx")
        doc.save(resume_file_path)
        
        self.results_text.insert(tk.END, f"\n✅ Generated {valid_entries} tailored resumes!\n")
        self.results_text.insert(tk.END, f"📄 Resume file saved: {resume_file_path}\n")
            
    def generate_single_resume(self, client, company, position, job_description, template_type='java'):
        """Generate a single tailored resume using OpenAI and template type."""
        
        # Read template content to use as base
        template_path = self.java_template_path if template_type == 'java' else self.csharp_template_path
        
        try:
            # Load template to extract existing content
            template_doc = Document(template_path)
            template_content = "\n".join([paragraph.text for paragraph in template_doc.paragraphs if paragraph.text.strip()])
        except:
            # Fallback to hardcoded content if template can't be read
            template_content = self.java_resume if template_type == 'java' else self.csharp_resume
        
        prompt = f"""You are the top resume writer in the world. Your job is to take the job description I provide and tailor a resume so that it is ATS-optimized, keyword-rich, and highly compelling.

Template Type Selected: {template_type.upper()}
Reason: Based on job description analysis, this position requires {template_type} expertise.

Follow these steps:

1. Keyword Extraction
   * Identify the most important hard skills, technical tools, industry terms, and role-specific keywords from the job description.
   * Highlight must-have ATS keywords that are critical for this role.

2. Resume Tailoring
   * Create a complete resume tailored for this specific {company} {position} role.
   * Use the XYZ method (Accomplished [X] as measured by [Y], by doing [Z]).
   * Incorporate identified keywords naturally throughout the resume.
   * Emphasize impact, metrics, and outcomes (not just duties).
   * Focus on {template_type}-related technologies and frameworks.
   * Do not include an objective statement.

3. Structure
   * SKILLS section with relevant technologies
   * WORK EXPERIENCE with 3-5 bullet points per role
   * EDUCATION section
   * PROJECTS section
   * Use professional formatting

Base Template Content (adapt and enhance):
{template_content}

Target Job Description:
Company: {company}
Position: {position}
Description: {job_description}

Generate a complete, ATS-optimized resume tailored specifically for this role."""

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "user", "content": prompt}
            ],
            max_tokens=4000,
            temperature=0.7
        )
        
        return response.choices[0].message.content
            
    def update_results(self, results):
        """Update the UI with processing results."""
        self.results_text.delete(1.0, tk.END)
        
        # Display results
        self.results_text.insert(tk.END, "🎉 PROCESSING COMPLETE!\n")
        self.results_text.insert(tk.END, "=" * 50 + "\n")
        self.results_text.insert(tk.END, f"📊 Total rows processed: {results['total_rows']}\n")
        self.results_text.insert(tk.END, f"✅ Valid entries: {results['valid_entries']}\n")
        self.results_text.insert(tk.END, f"⏭️ Skipped entries: {results['skipped_entries']}\n")
        self.results_text.insert(tk.END, f"📄 DOCX files created: {len(results['created_files'])}\n")
        self.results_text.insert(tk.END, f"📁 Output directory: {self.output_directory.get()}\n\n")
        
        if results['created_files']:
            self.results_text.insert(tk.END, "📋 Created files:\n")
            for filepath in results['created_files']:
                filename = os.path.basename(filepath)
                self.results_text.insert(tk.END, f"  • {filename}\n")
        
        self.status_label.config(text=f"✅ Success! Created {len(results['created_files'])} DOCX files")
        
    def show_error(self, message):
        """Show error message."""
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, f"❌ ERROR: {message}")
        self.status_label.config(text="❌ Processing failed")
        
    def processing_complete(self):
        """Called when processing is complete."""
        self.processing = False
        self.process_button.config(state="normal")


def main():
    """Main function to run the drag and drop application."""
    root = tkdnd.Tk()
    
    # Configure style
    style = ttk.Style()
    style.theme_use('clam')
    
    # Create and run the application
    app = DragDropApp(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()


if __name__ == "__main__":
    main()
