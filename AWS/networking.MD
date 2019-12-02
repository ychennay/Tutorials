# Networking

## Use Cases

Customer wants to use an existing VPC for one of their regulatory application workloads. VPC network traffic risks associated with the application instances communicating with unauthorized external hosts must be addressed. This new workload should not introduce any potential network bottlenecks or require application-level configuration changes due to compliance requirements.
Which architectural improvement would you recommend to meet company requirements?

**Proposal:** Configure an IDS/IPS (intrusion detection/prevention system) host-based solution that installs agents on the application instances to control outbound network traffic. Configure **NAT gateways** in each Availability Zone for egress traffic.

**Alternate solution:** Leverage forward proxy servers for web and URL content filtering and configure a Proxy ELB to distribute network traffic from the application instances to a fleet of proxy server instances for egress traffic.
