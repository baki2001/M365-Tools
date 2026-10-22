"""groups / onboard / offboard CLI via typer's CliRunner with the in-memory Graph fake."""

from __future__ import annotations

import json
import tomllib
from typing import Any

import pytest
from typer.testing import CliRunner

from m365tools import cli, services
from m365tools.core.config import AppConfig, add_tenant, save_config

from .fake_graph import FakeGraphGateway, make_fake, write_registry

ENV = {"COLUMNS": "200"}
runner = CliRunner()


@pytest.fixture
def gw(monkeypatch: pytest.MonkeyPatch) -> FakeGraphGateway:
    cfg = AppConfig()
    add_tenant(cfg, "contoso", tenant_id="contoso.onmicrosoft.com", client_id="cid",
               sharepoint_host="contoso.sharepoint.com")
    save_config(cfg)
    write_registry(services.groups_registry_path("contoso"))
    fake = make_fake()

    def factory(cfg: Any, tenant: Any = None, sink: Any = None, gateway: Any = None) -> Any:
        return services.make_groups_job(cfg, tenant, sink, gateway=fake)

    monkeypatch.setattr(cli, "make_groups_job", factory)
    return fake


def run(*args: str) -> Any:
    return runner.invoke(cli.app, list(args), env=ENV)


def test_inventory(gw: FakeGraphGateway) -> None:
    res = run("groups", "inventory")
    assert res.exit_code == 0, res.output
    assert "inventory summary" in res.output and "report.html" in res.output
    res = run("groups", "inventory", "--json", "--quiet")
    data = json.loads(res.stdout)
    assert data["people"] == 12 and data["groups"] == 8 and data["tenant"] == "contoso"


def test_drift_and_strict(gw: FakeGraphGateway) -> None:
    res = run("groups", "drift")
    assert res.exit_code == 0, res.output
    assert "unexplained" in res.output and "not_created" in res.output
    assert run("groups", "drift", "--strict").exit_code == 1
    data = json.loads(run("groups", "drift", "--json").stdout)
    assert {f["category"] for f in data["findings"]} >= {"missing", "unregistered", "unassigned"}


def test_suggest_accept_preview_then_apply(gw: FakeGraphGateway) -> None:
    path = services.groups_registry_path("contoso")
    before = path.read_text()
    res = run("groups", "suggest", "--accept", "3=Sales")
    assert res.exit_code == 0, res.output
    assert "Sales Manager" in res.output and "Preview only" in res.output and path.read_text() == before
    res = run("groups", "suggest", "--accept", "3=Sales", "--apply")
    assert res.exit_code == 0 and "registry change(s) written" in res.output
    assert tomllib.loads(path.read_text())["people"]["hugo@contoso.com"]["bundles"] == ["Sales"]
    assert run("groups", "suggest", "--accept", "42").exit_code == 1


def test_explain(gw: FakeGraphGateway) -> None:
    res = run("groups", "explain", "frank")
    assert res.exit_code == 0, res.output
    assert "missing" in res.output and "Servicedesk" in res.output
    data = json.loads(run("groups", "explain", "SG-Doc-Finance", "--json").stdout)
    assert data["type"] == "group" and data["unexplained"] == ["kim@contoso.com"]
    res = run("groups", "explain", "nobody")
    assert res.exit_code == 1 and "No person or group" in res.output


def test_sync_and_import(gw: FakeGraphGateway) -> None:
    res = run("groups", "sync")
    assert res.exit_code == 0 and "SG-App-Payroll" in res.output and "Preview only" in res.output
    assert gw.calls == []
    res = run("groups", "sync", "--apply")
    assert res.exit_code == 0, res.output
    assert gw.calls == [("create_group", "SG-App-Payroll", "Payroll app")]
    res = run("groups", "import", "--apply", "--json")
    assert json.loads(res.stdout)["groups"][0]["name"] == "SG-Doc-Archive"
    raw = tomllib.loads(services.groups_registry_path("contoso").read_text())
    assert "SG-Doc-Archive" in raw["groups"] and raw["groups"]["SG-App-Payroll"]["id"]


def test_onboard_preview_apply_and_check(gw: FakeGraphGateway) -> None:
    res = run("onboard", "jan@contoso.com", "--bundle", "Finance", "--addon", "SG-Doc-Sales")
    assert res.exit_code == 0, res.output
    assert "planned" in res.output and "App:Payroll" in res.output and gw.calls == []
    res = run("onboard", "jan@contoso.com", "--like", "dana@contoso.com", "--apply", "--json")
    assert res.exit_code == 0, res.output
    data = json.loads(res.stdout)
    assert data["applied"] and {c["status"] for c in data["plan"]["changes"]} == {"done"}
    assert run("onboard", "check", "jan@contoso.com").exit_code == 0
    res = run("onboard", "check", "frank@contoso.com")
    assert res.exit_code == 1 and "SG-App-Jira" in res.output
    assert run("onboard", "jan@contoso.com", "--bundle", "Nope").exit_code == 1
    assert run("onboard", "a", "b").exit_code == 1


def test_offboard(gw: FakeGraphGateway) -> None:
    res = run("offboard", "anna@contoso.com")
    assert res.exit_code == 0 and "would be removed" in res.output and gw.calls == []
    res = run("offboard", "anna@contoso.com", "--apply")
    assert res.exit_code == 0, res.output
    assert "not disabled" in res.output
    assert "anna@contoso.com" not in gw.members_of("SG-Doc-Hub")
    raw = tomllib.loads(services.groups_registry_path("contoso").read_text())
    assert raw["people"]["anna@contoso.com"]["offboarded"]


def test_failed_write_exits_1(gw: FakeGraphGateway) -> None:
    gw.fail_on.add("SG-App-Jira")
    res = run("onboard", "frank@contoso.com", "--apply")
    assert res.exit_code == 1 and "failed" in res.output


def test_unknown_tenant(gw: FakeGraphGateway) -> None:
    res = run("groups", "drift", "--tenant", "nope")
    assert res.exit_code == 1


def test_setup_help_lists_group_permissions() -> None:
    res = run("tenant", "setup-help")
    for perm in ("Group.Read.All", "User.Read.All", "Directory.Read.All", "Application.Read.All",
                 "Group.ReadWrite.All", "GroupMember.ReadWrite.All"):
        assert perm in res.output
