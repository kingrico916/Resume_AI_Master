"""
CalCareers Application Packager - Checklist Generator
Generates human-executable VSC checklists per CalCareers process.
"""

from typing import Dict
from core.data_models import PackageInput
from datetime import datetime


class ChecklistGenerator:
    """Generates standardized VSC checklists for CalCareers applications."""
    
    # CalCareers Apply-Now flow reference
    CALCAREERS_PROCESS_URL = "https://calcareers.ca.gov/CalHRPublic/GeneralInfo/FAQ.aspx"
    
    def __init__(self):
        pass
    
    def generate_checklist(self, pkg_input: PackageInput, decisions: Dict) -> str:
        """
        Generate complete VSC checklist.
        
        Args:
            pkg_input: Package input data
            decisions: Decisions dict from DecisionEngine
            
        Returns:
            Formatted checklist as string
        """
        sections = []
        
        # Header
        sections.append(self._generate_header(pkg_input))
        
        # Job basics
        sections.append(self._generate_job_basics(pkg_input))
        
        # Template selection
        sections.append(self._generate_template_section(decisions))
        
        # Application process
        sections.append(self._generate_process_section())
        
        # Required documents
        sections.append(self._generate_required_docs_section(decisions))
        
        # Optional documents
        sections.append(self._generate_optional_docs_section(decisions))
        
        # Veterans' Preference (if applicable)
        if pkg_input.veterans_preference.claiming_veterans_preference:
            sections.append(self._generate_vp_section(pkg_input, decisions))
        
        # Final submission checklist
        sections.append(self._generate_final_checklist())
        
        return "\n\n".join(sections)
    
    def _generate_header(self, pkg_input: PackageInput) -> str:
        """Generate checklist header."""
        return f"""{'='*80}
CALCAREERS APPLICATION CHECKLIST
{'='*80}
Candidate: {pkg_input.candidate.legal_name}
Package: {pkg_input.get_package_name()}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
{'='*80}"""
    
    def _generate_job_basics(self, pkg_input: PackageInput) -> str:
        """Generate job basics section."""
        job = pkg_input.job
        deadline = job.final_filing_date if job.final_filing_date else "CHECK POSTING"
        department = job.department if job.department else "See posting"
        
        return f"""JOB TARGET
----------
JC Number: {job.jc_number}
Classification: {job.classification_title}
Department: {department}
Final Filing Date: {deadline}

ACTION: Verify all details on CalCareers posting before applying."""
    
    def _generate_template_section(self, decisions: Dict) -> str:
        """Generate STD 678 template section."""
        track = decisions['template_track']
        
        return f"""STD 678 TEMPLATE
----------------
Selected Track: {track}

INSTRUCTIONS:
1. Use the {track} template from CalCareers STD 678 library
2. Complete all required fields (no blank required fields)
3. Tailor "Relevant Experience" section to match job duties
4. Review for accuracy - this becomes your official application

FILE LOCATION: 02_STD678/"""
    
    def _generate_process_section(self) -> str:
        """Generate application process section."""
        return f"""APPLICATION PROCESS (CalCareers Apply-Now Flow)
-----------------------------------------------
Reference: {self.CALCAREERS_PROCESS_URL}

STEP 1: Navigate to job posting on CalCareers
[ ] Search by JC number or classification
[ ] Verify posting matches target job

STEP 2: Create/login to CalCareers account
[ ] Use candidate's email for account
[ ] Save login credentials securely

STEP 3: Complete online application
[ ] Upload STD 678 from package
[ ] Upload required documents
[ ] Upload optional documents (recommended)
[ ] Review all entries for accuracy

STEP 4: Review and certify
[ ] Review entire application
[ ] Check all uploaded documents
[ ] Certify information is accurate
[ ] NOTE: VSC must NOT click final "Submit" - candidate must certify personally

STEP 5: Candidate final submission
[ ] Candidate reviews application
[ ] Candidate certifies under penalty of perjury
[ ] Candidate clicks "Submit Application"
[ ] Save confirmation email to 06_Receipts/"""
    
    def _generate_required_docs_section(self, decisions: Dict) -> str:
        """Generate required documents section."""
        docs = decisions['base_required_docs']
        doc_list = "\n".join([f"[ ] {doc}" for doc in docs])
        
        return f"""REQUIRED DOCUMENTS
------------------
{doc_list}

FILE LOCATION: 03_Required_Docs/

ACTION: Ensure all required documents are prepared and ready to upload."""
    
    def _generate_optional_docs_section(self, decisions: Dict) -> str:
        """Generate optional documents section."""
        docs = decisions.get('optional_docs', [])
        doc_list = "\n".join([f"[ ] {doc}" for doc in docs])
        
        return f"""OPTIONAL DOCUMENTS (Recommended)
--------------------------------
{doc_list}

FILE LOCATION: 04_Optional_Docs/

NOTE: While optional, these documents strengthen the application and are 
recommended for competitive positions."""
    
    def _generate_vp_section(self, pkg_input: PackageInput, decisions: Dict) -> str:
        """Generate Veterans' Preference section."""
        basis = decisions['vp_basis']
        sections = ", ".join([str(s) for s in decisions['required_sections']])
        attachments = decisions['required_attachments']
        att_list = "\n".join([f"[ ] {att}" for att in attachments])
        address = decisions['submission_address']
        
        return f"""VETERANS' PREFERENCE
--------------------
Authority: CalHR-1093 (https://www.calhr.ca.gov/Documents/CalHR-1093.pdf)
Legal Authority: California Government Code §18973.1

BASIS: {basis}
REQUIRED CALHR 1093 SECTIONS: {sections}

REQUIRED ATTACHMENTS:
{att_list}

FILE LOCATION: 05_VeteransPreference/

CALHR 1093 SUBMISSION
---------------------
Complete CalHR 1093 form and mail to:

{address}

IMPORTANT:
- CalHR 1093 must be submitted separately to CalHR
- Include all required attachments listed above
- Keep copy of mailed package for records
- Mailing deadline may differ from application deadline - verify on form"""
    
    def _generate_final_checklist(self) -> str:
        """Generate final pre-submission checklist."""
        return f"""FINAL PRE-SUBMISSION CHECKLIST
------------------------------
[ ] All required documents uploaded to CalCareers
[ ] All optional documents uploaded (if applicable)
[ ] STD 678 reviewed for accuracy and completeness
[ ] Veterans' Preference documents mailed to CalHR (if applicable)
[ ] Candidate has reviewed entire application
[ ] Candidate is ready to certify and submit
[ ] Confirmation email will be saved to 06_Receipts/

{'='*80}
END OF CHECKLIST
{'='*80}"""
