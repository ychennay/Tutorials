# Databases

## Important Metrics

* **RPO (recovery point objective)**: maximum period of acceptable data loss (*how much data are you okay losing in the event of a failure?*)

* **RTO (recovery time objective)**: maximum acceptable time to recover data and resume normal operations after failure.

Replication to create read-replicas is **asynchronous**. However, for Multi-AZ deployments, replication between the primary and standby instance is **synchronous**.

## Point in Time Recovery

Enabling **point-in-time recovery** generates database change logs written to S3 each 5 minutes. 

## Provisioned IOPs

- use for any production application that requires consistent I/O, particularly OLTP workloads

## Service Levels

AWS gaurantees service levels for provisioned IOPS within a **10%** range for 99.9% of the year (ie., any time except for 2 hours and 45 minutes a year).

If you need to perform custom service configurations on a database engine, then you are better off deploying AWS EC2 instances and spinning up the database on the instances (since you'll need to SSH in for custom work).

Only MySQL and MariaDB database instances support creating read replicas of read replicas.

The maximum size of an RDS DB instance's storage capacity is 6TB (underlying EBS volume).

AWS RDS instance storage sizes can only be **scaled up, not down**.

Every DB instance has a weekly maintenance window (30 minutes long).

Events that can force the RDS instance to be take offline:
- security or operating system patching
- scaling compute operations and class instances
- major database version upgrades (also cannot revert back to previous version)

Users can elect when to upgrade to **major version releases and most minor upgrades**.

DUring an upgrade for storage size, the following is performed:
1. Performance maintenance on standby
2. Failover to standby (now that it is upgraded)
3. Perform maintenance on former primary

### Failovers for Multi-AZ

In Multi-AZ deployments, there is always a **primary** and **standby** DB instance. Read/writes go to the primary instance. Automated backups and snapshots are performed against the standby RDS instance.

During failover, the DNS record of the RDS endpoint is updated to point to the IP address of the standby instance. But the endpoint does not change.

To initiate a manual failover, initiate a reboot with a failover on the primary RDS instance. When an automatic failover occurs, you can **subscribe to an SNS topic** to receive a notification for an automatic failover event. 

The **InnoDB** database engine supports RDS read replicas. The new DB instance promoted from a read-replica to primary will inherit the following:
* backup window (but NOT the maintenance window)
* backup retention period
* database parameter group

The other, non-promoted read replicas will continue to function as before, replicating from the original primary DB instance.

## General-Purpose SSD

For each GB of data allocated for storage, RDS will provide 3 IOPs, up to 10,000 IOPS:

* 20GB $\rightarrow$ 60 IOPs
* * 100GB $\rightarrow$ 300 IOPs

Maximum throughput for gp2 storage type is 1,280 Mbps (160 MBps). Suppose you have MariaDB with **page size of 16KB (.128 Mbs)**:

$$
1280Mbps\times\frac{1 page}{0.128 Mbs}\times\frac{1IO}{1 page} = 10000 IOPs
$$

You'll also need to account for how much storage to allocate

$$
10000 IOPS \times \frac{1GB}{3 IOPS} = 3333.3 GB
$$

## Credits
You can also use **burst durations** if you might need high IOPs during certain occasions but not much storage. For volumes under 1TB, you can temporarily burst to 3,000 IOPS.
$$
B = \frac{C}{3000 - 3S}
$$
$B$ is the burst duration in seconds, $C$ is the credit balance, and $S$ is the storage size in GB. A database instance is credited with an initial balance of 5,400,000 IOPs. For instance, if you have a 300GB volume and a completely full original credit balance,

$$
B = \frac{5,400,000}{3000 - 3\times300} = 2571
$$
Your database could burst for about 2,571 seconds, or roughly 40 minutes. Credit balances are replenished at a rate of 1 IOPS per second per baseline IOPS. If you have a 200GB volume database, your baseline IOPS is 600 IOPS, so to refill your credits up to the maximum of 5,400,000 credits,

$$
5400000 \times \frac{1 second}{600 IOPS}\times \frac{1 hour}{3600 seconds} = 2.5 hours
$$

## Provisioned IOPs

No concept of bursting for **io1 storage**.

The ratio of storage in gigabytes to provisioned IOPS must be at least **50:1**. For instance, 20,000 IOPs will mean at least 400GB storage.

You can allocate between 1,000 to 40,000 provisioned IOPs and between 100GB to 16TB of data. 

# DynamoDB 

NoSQL databases such as DynamoDB are ideal when you will query data based on only one attribute (column), as you can specify this as the primary key and optimize w/ an index. 

### Integration with SQS

1. Application writes to middle-tier SQS app, which writes to DynamoDB
2. If DynamoDB is throttled, then write to SQS first as a buffer queue.
3. An SQS queue draining app checks the queue and drains it with writes into DynamoDB when queue length > 0.

### DynamoDB Global Tables

Use cases:
* Use Global tables as a **cache layer** (application in geo specific region accesses closest Global Table for resiliency, lower latency).
* 

### DynamoDB Accelerator (DAX)

### DynamoDB Streams

DynamoDB Streams captures a **time-ordered sequence of item-level modifications** in any DynamoDB table and stores this information in a log for up to 24 hours. Applications can access this log and view the data items as they appeared before and after they were modified, in near-real time.

![Token](public/StreamsAndTriggers.png)

Sample use cases:
* An application in one AWS Region modifies the data in a DynamoDB table. A second application in another Region reads these data modifications and writes the data to another table, creating a replica that stays in sync with the original table.

Can be used instead of creating a worker-tier application and overhead of syncing across multi-region tables. Should typically not try to write data to a DynamoDB table in each region directly.

### Strongly Consistent versus Eventually Consistent

A **single strongly consistent** read of 4KB per second will consume 1 read capacity unit. An **eventual consistent** read of 2KB per second will consume 1 RCU (half the cost in terms of RCUs as strongly consistent reads).

### Autoscaling

1. Specify a maximum and minimum **Read** and **Write** Capacity Unit (RCU and WCUs), along with **utilization**.

* RCU: min 10, max 50
* utilization goal: 50%
* actually consume 20 -> autoscaling will set RCU at 40.

Items are distributed across partitions based on the **partition key** (AKA **hask key**)

### Composite Primary Keys

Use the combination of a **partition key** and a **sort key** to unique identify each item. Sort keys are ideal for non-unique columns that are commonly used to order records, like the `order_date` field.

## Redshift

Redshift is an **OLAP** database with **columnar storage** ideal for analytical or complex relational queries.

Types of nodes:
* **Dense compute nodes**: can store up to 326TB of data on magnetic storage.
* **Dense storage nodes**: can store up to 2PB of data on SSD (use this if you want to optimize read speeds)

Any cluster with more than one compute node includes a **leader node**.

Distribution styles:
* **EVEN**: leader node spreads data evenly across all compute notes (default style)
* **KEY**: data is distributed according to the value in a single column
* **ALL**: all tables are distirbuted to all compute nodes.

## Aurora

A **regional service** database engine compatible with MySQL (5x throughout) and PostgreSQL (3x throoughput), where compute and storage are partitioned, with a primary instance that allows reads and writes. Sync with read replicas is **synchronous**.

Two types of DB instances:
* **primary**: supports reads/writes, and is a singleton instance (only 1 primary instance)
* **replica**: also connects to the same underlying storage volume, supports only read operations (a max of 15 replica instances per Aurora DB cluster)

THe underlying storage (**cluster volume** can scale up to **64TB**) is SSD and stored on **3 AZs in the same region**. Each availability zone that the cluster volume resides in will have 2 copies of the DB cluster data. Storage replication is **independent of number of instances**. Table size is limited to 64TB.

No concept of a **standby instance**.

### Differences with Amazon RDS
In RDS,
* 1 **primary instance** in **AZ1** that handles all writes (and some reads potentially). If the primary instance reboots, all the replica instances will also be rebooted.
* **Read replicas** in other Availability Zones that asynchronously replicated from the master instance. AWS recommends that the replicas are the same DB instance class as the primary instance.

### Migration from Amazon RDS
1. Create an RDS snapshot, and then restore snapshot in Aurora
2. Set up one-way replication.

### Security

IAM can be used for identity access management, and **VPC endpoint** to route traffic from a private VPC without it crossing the public internet. Use a VPC security group on the VPC endpoint to control which instances in your VPC get access to the endpoint. 

Automated backups, read replicas, snapshots can all be encrypted via AES-256. You cannot convert an unencrypted DB cluster to encrypted one easily:
1. Take a snapshot of the unencrypted DB cluster.
2. Specify a KMS encryption key upon restore.

You cannot go from encrypted $\rightarrow$ unecrypted cluster, however. 

### Point in Time Recovery

- automated backups and DB snapshots for MySQL is supported for only InnoDB.
- uses RDS automated backups and transaction logs
- if backup window time is set to 0, then no backup snapshots are created, and thus no ability for PiTR.
- if the point in time to restore to is within the retention period, you can specify up to a specific second.
- the most recent backup snapshot is retrieved, and transaction logs are pplied to restore your DB to that point in time.

### Autoscaling

Aurora is an example of **single master replication**, since there is only one master instance, autoscaling for Aurora applies only to read replicas. Autoscaling must start with at least one replica. 

Note that Aurora replication **replicates entire tables**, not individual records.

For **primary failover**: if the primary instance goes out of business, then Aurora will check if there is a read replica. If there is, it will promote this replica to primary by pointing the **cluster endpoint** to the new read replica instance. If there are no replicas, it will create a new primary instance, but this means that your Aurora cluster will be unavailable for the amount of time it takes for Aurora to spin up the new instance. 

If an application consuming the Aurora endpoints is caching the IP addresses of the primary instance, set TTL time to a low number in case of failover (IP will point to something )

### Connections and Endpoints

A **handler endpoint** is a URL (host name and port) that is provided by Aurora that contains logic to determine where to route a request for a query, or a write transaction, etc. to the IP address of the DB cluster instance to handle the request, providing failover support for the cluster. This endpoint cannot be modified. 

A **reader endpoint** is a URL that connects to one of the available read replicas in the DB cluster, providing load-balancing for read-only connections. 

An **instance endpoint** is a URL for a specific DB instance, used to provided direct, granular control over individual connections or instances in the cluster. You might use this to configure **fine-grained load balancing for a specific request**. 

A **custom endpoint** (up to 5, not available for Aurora Serverless) can be created by users specific to a particular workload. From the documentation:

> For example, you may provision a set of Aurora Replicas to use an instance type with higher memory capacity in order to run an analytics workload. A custom endpoint can then help you route the analytics workload to these appropriately-configured instances, while keeping other instances in your cluster isolated from this workload. 

### Sharing with Other Accounts

You can share the Aurora snapshots with at most 20 other accounts. If the snapshot is encrypted, make sure it is not encrypted with a default KMS key, but rather a custom key that you gran access to the `Decrypt` action to the shared account. The other account will need to take the snapshot and spin up a cluster from it. 

### Aurora Global Database

Consists of one primary AWS region and one secondary (read-only) region. AWS Aurora will support up to 5 cross regions with DB cluster read replicas.  You can also promote a single read replica instance into its own standalone cluster.

Typical use cases:
* Disaster recovery in a different region (promote replica in different region to primary)
* You start in North America with your business customers but now many of your queries are originating from APAC- you **create a secondary AWS Region** to achieve lower latency reads

### Integratins with other AWS Services

- Use Aurora MySQL DB stored procedures and functions to invoke AWS Lambda
- Load data into Aurora table from S3 bucket (MySQL)
- Query data from MySQL DB cluster and save directly as text files in S3 bucket

### Aurora Serverless

AWS will manage a **DB capacity warm pool** of instances that can be quickly provisioned on demand. 

Use case:
* intermittent, unpredictable workloads
* low-volume blog site
* new applications where the performance requirements are unknown and need to be discovered
* more cost-effective in this case
* architect doesn't know best instance class size to use
* development and test databases

Currently, failover in Serverless takes longer than provisioned cluster, as the DB cluster is created in a single availability zone, and Aurora must recreate the DB instance in a different AZ (**automatic multi-AZ failover**).

## Use Cases

A company has a **24 TB MySQL database** in its on-premises data center that grows at the rate of 10 GB per day. The data center is connected to the company’s AWS infrastructure with a 50 Mbps VPN connection. The company is migrating the application and workload to AWS. The application code is already installed and tested on Amazon EC2. The company now needs to migrate the database and wants to go live on AWS within 3 weeks.

Which of the following approaches meets the schedule with LEAST downtime?

Proposal:
1. Create a database export locally and import into AWS Snowball
2. Launch Aurora DB instance (MySQL RDS only supports a capacity of 16TB)
3. Load data into the Aurora DB via the Snowball export
4. Set up database replication from on-premise DB to RDS Aurora DB via the VPN.
5. Switch over DNS CNAME/A record to the Aurora instance DNS endpoint
6. Stop replication once it is in sync.