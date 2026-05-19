"""
CalCareers Application Packager - Data Models
Defines the data contract for candidate, job, and veterans' preference inputs.
"""

from dataclasses import dataclass, field
from typing import Optional, List, Literal
from datetime import date, datetime
from enum import Enum


class TemplateTrack(Enum):
    """STD 678 template tracks supported by CalCareers."""
    ANALYST = "ANALYST"
    IT = "IT"
    OPS = "OPS"


class ApplicationBasis(Enum):
    """Veterans' Preference basis per CalHR 1093."""
    SELF_VETERAN = "SELF_VETERAN"
    SELF_ACTIVE_MILITARY = "SELF_ACTIVE_MILITARY"
    SPOUSE_DECEASED_VETERAN = "SPOUSE_DECEASED_VETERAN"
    SPOUSE_100_DISABLED_VETERAN = "SPOUSE_100_DISABLED_VETERAN"


@dataclass
class EducationEntry:
    """Educational background entry."""
    institution: str
    degree: str
    major: str
    graduation_date: Optional[str] = None
    gpa: Optional[str] = None


@dataclass
class WorkExperienceEntry:
    """Work experience entry."""
    employer: str
    job_title: str
    start_date: str
    end_date: Optional[str]
    duties: str
    supervisor: Optional[str] = None
    phone: Optional[str] = None


@dataclass
class CandidateProfile:
    """Complete candidate profile per data contract."""
    legal_name: str
    date_of_birth: str  # YYYY-MM-DD
    address: str
    phone: str
    email: str
    education_entries: List[EducationEntry]
    work_experience_entries: List[WorkExperienceEntry]
    ecos_id: Optional[str] = None


@dataclass
class JobTarget:
    """Target CalCareers job posting."""
    jc_number: str
    classification_title: str
    department: Optional[str] = None
    final_filing_date: Optional[str] = None
    required_documents: List[str] = field(default_factory=list)


@dataclass
class VeteransPreference:
    """Veterans' Preference claim per CalHR 1093."""
    claiming_veterans_preference: bool
    application_basis: Optional[ApplicationBasis] = None
    
    def get_required_sections(self) -> List[int]:
        """Returns required CalHR 1093 sections based on basis."""
        if not self.claiming_veterans_preference or not self.application_basis:
            return []
        
        # All bases require sections 1, 2, 4
        base_sections = [1, 2, 4]
        
        # Spouse 100% disabled requires section 3
        if self.application_basis == ApplicationBasis.SPOUSE_100_DISABLED_VETERAN:
            return [1, 2, 3, 4]
        
        return base_sections
    
    def get_required_attachments(self) -> List[str]:
        """Returns required attachments based on basis per CalHR 1093."""
        if not self.claiming_veterans_preference or not self.application_basis:
            return []
        
        attachments = []
        basis = self.application_basis
        
        # DD214 requirements
        if basis in [ApplicationBasis.SELF_VETERAN, 
                     ApplicationBasis.SPOUSE_DECEASED_VETERAN,
                     ApplicationBasis.SPOUSE_100_DISABLED_VETERAN]:
            attachments.append("DD214")
        
        # CalHR 1094 for active military
        if basis == ApplicationBasis.SELF_ACTIVE_MILITARY:
            attachments.append("CalHR 1094")
        
        # Marriage certificate for spouse-based
        if basis in [ApplicationBasis.SPOUSE_DECEASED_VETERAN,
                     ApplicationBasis.SPOUSE_100_DISABLED_VETERAN]:
            attachments.append("Marriage Certificate")
        
        # Death certificate
        if basis == ApplicationBasis.SPOUSE_DECEASED_VETERAN:
            attachments.append("Death Certificate")
        
        # Disability award letter
        if basis == ApplicationBasis.SPOUSE_100_DISABLED_VETERAN:
            attachments.append("Disability Award Letter")
        
        return attachments


@dataclass
class PackageInput:
    """Complete input package for automation."""
    candidate: CandidateProfile
    job: JobTarget
    template_track: TemplateTrack
    veterans_preference: VeteransPreference
    
    def get_package_name(self) -> str:
        """Generate standardized package folder name."""
        last_name = self.candidate.legal_name.split()[-1]
        first_name = self.candidate.legal_name.split()[0]
        jc = self.job.jc_number.replace(" ", "")
        return f"{last_name}_{first_name}_JC{jc}_CalCareersPackage"
    
    def get_filename_prefix(self) -> str:
        """Generate standardized filename prefix."""
        last_name = self.candidate.legal_name.split()[-1]
        first_name = self.candidate.legal_name.split()[0]
        jc = self.job.jc_number.replace(" ", "")
        return f"{last_name}_{first_name}_JC{jc}"


@dataclass
class AuditRecord:
    """Audit log entry for package generation."""
    timestamp: datetime
    package_name: str
    inputs_received: dict
    decisions_made: dict
    governing_authorities: List[str]
    missing_data: List[str]
    
    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            "timestamp": self.timestamp.isoformat(),
            "package_name": self.package_name,
            "inputs_received": self.inputs_received,
            "decisions_made": self.decisions_made,
            "governing_authorities": self.governing_authorities,
            "missing_data": self.missing_data
        }
