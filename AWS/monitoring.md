# Monitoring

## CloudTrail

An **event** in the context of CloudTrail is an action (`ListObjects`) a principal (user `ychen`) performs on an AWS resource (S3 bucket). 

A **trail** is a setup that delivers event logs to a user-specified S3 bucket. A trail can be configured for either a specific region, or all regions.

By default, CloudTrail logs **90 days of management events**, with separate vent histories for each region. Global events like Route 53 or IAM are included in each region's event history. If you want more than 90 days of event history or want to customize event types, you should create a **trail** (up to 5 trails per region). A trail that is global will count as one trail per region.

Creating a trail via the management console defaults to **all regions**. On the API or CLI, the default is a **single region.**

### Management Events

Management events are grouped into two types:
* **read-only** events (`DescribeInstances` API call on EC2, any action that reads a resource but cannot modify it)
* **write-only** events (API operations such as logging into management console)

### Data Events

### Log File Integrity

### Sample Use Cases

#### On-Premise Monitoring Integration
Your customer’s **existing on-premises monitoring system** currently receives performance notifications from local management systems and data from agents installed in on-premises operating systems. The customer wants to use this on-premises monitoring system to **monitor Amazon EC2 instance performance and state changes of EC2 instances**.

What should you do to **monitor EC2 instance state changes and performance**?

Proposed:
1. Configure a **CloudWatch Events rule** for Amazon EC2 instance state changes. 
2. Have the Events rule **trigger an AWS Lambda function** that notifies your monitoring system. 
3. **Pull CloudWatch metric data** into your monitoring system.

Considerations:
- CloudTrail is for API calls, and is not a relevant service for this problem (we are worried about performance monitoring and state transitions, not user API calls)