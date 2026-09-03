"""Rule registry. Importing this package registers every detector."""
from .base import (CRITICAL, INFO, WARNING, Finding, Fix, RuleMeta, addr,
                   is_currency_format, registry, rule, shape_of, to_r1c1)
from . import correctness, hygiene, structural  # noqa: F401  (import for side effects)

__all__ = ["Finding", "Fix", "RuleMeta", "registry", "rule", "to_r1c1", "shape_of",
           "addr", "is_currency_format", "CRITICAL", "WARNING", "INFO"]
