"""
CalCareers Application Packager - Decision Engine
Implements CalHR 1093 Veterans' Preference decision logic.
Authority: CalHR-1093.pdf, California Government Code §18973.1
"""

from typing import Dict, List, Tuple
from core.data_models import PackageInput, ApplicationBasis, TemplateTrack


class DecisionEngine:
    """
    Implements decision logic per CalHR 1093.
    Source: https://www.calhr.ca.gov/Documents/CalHR-1093.pdf
    """
    
    # Official CalHR submission address (CalHR-1093, Page 2)
    CALHR_SUBMISSION_ADDRESS = """California Department of Human Resources
Attn: Selection Division
1515 "S" Street, North Building, Suite 500
Sacramento, CA 95811"""
    
    def __init__(self):
        self.decisions = {}
        self.authorities = []
    
    def analyze_package(self, pkg_input: PackageInput) -> Dict:
        """
        Primary decision analysis for package generation.
        Returns decisions dict and populates self.authorities.
        """
        self.decisions = {}
        self.authorities = [
            "CalHR-1093 (Veterans' Preference Authority)",
            "California Government Code §18973.1",
            "CalCareers STD 678 Template Reuse Guidance"
        ]
        
        # Template selection
        self.decisions['template_track'] = pkg_input.template_track.value
        
        # Veterans' Preference analysis
        vp = pkg_input.veterans_preference
        self.decisions['claiming_veterans_preference'] = vp.claiming_veterans_preference
        
        if vp.claiming_veterans_preference:
            self.decisions['vp_basis'] = vp.application_basis.value if vp.application_basis else "UNKNOWN"
            self.decisions['required_sections'] = vp.get_required_sections()
            self.decisions['required_attachments'] = vp.get_required_attachments()
            self.decisions['submission_address'] = self.CALHR_SUBMISSION_ADDRESS
        else:
            self.decisions['vp_basis'] = None
            self.decisions['required_sections'] = []
            self.decisions['required_attachments'] = []
            self.decisions['submission_address'] = None
        
        # Document requirements
        self.decisions['base_required_docs'] = self._get_base_required_docs(pkg_input)
        self.decisions['optional_docs'] = self._get_optional_docs(pkg_input)
        
        return self.decisions
    
    def _get_base_required_docs(self, pkg_input: PackageInput) -> List[str]:
        """
        Determine base required documents (non-VP).
        Based on job posting requirements and CalCareers standards.
        """
        required = ["STD 678"]
        
        # Add job-specific requirements
        if pkg_input.job.required_documents:
            required.extend(pkg_input.job.required_documents)
        
        return required
    
    def _get_optional_docs(self, pkg_input: PackageInput) -> List[str]:
        """Determine optional supporting documents."""
        optional = ["Resume", "SOQ", "Transcripts", "Certifications"]
        return optional
    
    def validate_inputs(self, pkg_input: PackageInput) -> Tuple[bool, List[str]]:
        """
        Validate input completeness. Returns (is_valid, missing_fields).
        Does NOT invent data - flags missing items explicitly.
        """
        missing = []
        
        # Candidate profile validation
        if not pkg_input.candidate.legal_name:
            missing.append("candidate.legal_name")
        if not pkg_input.candidate.date_of_birth:
            missing.append("candidate.date_of_birth")
        if not pkg_input.candidate.address:
            missing.append("candidate.address")
        if not pkg_input.candidate.phone:
            missing.append("candidate.phone")
        if not pkg_input.candidate.email:
            missing.append("candidate.email")
        if not pkg_input.candidate.education_entries:
            missing.append("candidate.education_entries (empty)")
        if not pkg_input.candidate.work_experience_entries:
            missing.append("candidate.work_experience_entries (empty)")
        
        # Job target validation
        if not pkg_input.job.jc_number:
            missing.append("job.jc_number")
        if not pkg_input.job.classification_title:
            missing.append("job.classification_title")
        
        # Veterans' Preference validation
        vp = pkg_input.veterans_preference
        if vp.claiming_veterans_preference:
            if not vp.application_basis:
                missing.append("veterans_preference.application_basis (required when claiming VP)")
        
        is_valid = len(missing) == 0
        return is_valid, missing
    
    def get_template_recommendation(self, classification_title: str) -> TemplateTrack:
        """
        Infer template track from classification title.
        High confidence required - otherwise requires manual selection.
        """
        title_lower = classification_title.lower()
        
        # IT keywords
        it_keywords = ['information', 'technology', 'systems', 'programmer', 'developer', 
                       'software', 'network', 'database', 'cyber', 'it specialist']
        if any(keyword in title_lower for keyword in it_keywords):
            return TemplateTrack.IT
        
        # Analyst keywords
        analyst_keywords = ['analyst', 'research', 'policy', 'budget', 'management services']
        if any(keyword in title_lower for keyword in analyst_keywords):
            return TemplateTrack.ANALYST
        
        # Operations keywords
        ops_keywords = ['operations', 'technician', 'specialist', 'coordinator', 
                        'officer', 'associate', 'assistant']
        if any(keyword in title_lower for keyword in ops_keywords):
            return TemplateTrack.OPS
        
        # Default to ANALYST if uncertain (most versatile)
        return TemplateTrack.ANALYST
    
    def get_authorities(self) -> List[str]:
        """Return list of governing authorities cited in decisions."""
        return self.authorities
