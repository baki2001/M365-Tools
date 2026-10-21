"""The seven sections of the workspace, in sidebar order."""

from .base import Section
from .checks import ChecksSection
from .job import JobSection
from .lock import LockSection
from .plan import PlanSection
from .run import RunSection
from .runs import RunsSection
from .verify import VerifySection

SECTIONS: tuple[type[Section], ...] = (JobSection, ChecksSection, PlanSection, RunSection, VerifySection,
                                       LockSection, RunsSection)

__all__ = ["SECTIONS", "ChecksSection", "JobSection", "LockSection", "PlanSection", "RunSection",
           "RunsSection", "Section", "VerifySection"]
