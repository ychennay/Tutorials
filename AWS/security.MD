# Security

### Use Cases
Customer migrated their public website to AWS after suffering a major data center failure. Due to the urgency of the move, the customer’s operations team did not implement a load balancer or an Auto Scaling group. The **two web server instances have Elastic IP addresses** and their DNS provider returns both IPs when queried. This morning, the website was the target of malicious login attempts from about 10 IP addresses.

Which solution best offers quick and easy protection for the site from these malicious login attempts?

**Proposal:** Create a **DENY network access control list (ACL) for the known malicious IP addresses**, and associate it with the subnets containing the web server instances.

Considerations:
- Implementing AWS WAF tier w/ load balancers (WAF sandwich) is correct, but takes too long
- Moving DNS provider is not a quick operation
- Creating a new security group w/ predefined whitelisted IP addresses is doable but you'd need to add in every single allowed IP CIDR block