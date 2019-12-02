# Cost Control

## Best Practices

- **limit resource provisioning** using Service Control Policies (SCP) - follow policy of least permissions, and turn off services you know you are not using
- **tag everything**: extremely useful for audit compliance, cost reduction, and integration w/ **AWS Inspector**. Require tags in order to instantiate resources. If your organization is scaled large enough, consider revoking access to the Management console so that all deployments go through a CI/CD tooling pipeline.
- **Configure CloudWatch alarms** for billing thresholds.