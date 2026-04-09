from dataclasses import dataclass


@dataclass(frozen=True)
class Profile:
    name: str
    base_url: str
    jwt_url: str


PROFILES = {
    "dev": Profile(
        name="dev",
        base_url="https://iam.torknd-customer.dev.pd15.digitalgov.mtp",
        jwt_url="https://iam.torknd-customer.dev.pd15.digitalgov.mtp/jwt/",
    ),
    "psi": Profile(
        name="psi",
        base_url="https://pgs-psi-inner.digitalgov-torknd-psi-common.apps.k8s.prod1.pd40.sol.mtp",
        jwt_url="https://psi.pgs.gosuslugi.ru/getDebug",
    ),
    "prod": Profile(
        name="prod",
        base_url="http://pgs-prod-inner.digitalgov-torknd-prod1-common.apps.k8s.prod1.pd40.sol.mtp",
        jwt_url="https://pgs.gosuslugi.ru/getDebug",
    ),
}
