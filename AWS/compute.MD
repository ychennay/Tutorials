# Compute

## Reserved Instances

Running a reserved instance for less than full utilization is more or less wasting money (since you are getting billed for the full 24 hours per day).

## Scheduled Reserved Instances

Scheduled Reserved Instances (Scheduled Instances) enable you to purchase capacity reservations that recur on a daily, weekly, or monthly basis, with a specified start time and duration, for a one-year term. You reserve the capacity in advance, so that you know it is available when you need it. You pay for the time that the instances are scheduled, even if you do not use them.

# EMR

## Node Types

Three types of nodes:

### Master
Manages the cluster and typically runs master components of distributed applications. Ideal for reserved/on-demand instances.

The master instance group in an EMR cluster always consists of a single node or three master nodes, so it can't scale after you initially configure it. You work with the core instance groups and task instance groups to scale out and scale in a cluster. It's possible to have a cluster with only a master node, and no core or task nodes. You must have at least one core node at cluster creation in order to scale the cluster.

### Core
Managed by the master node. Core nodes run the Data Node daemon to coordinate data storage as part of the Hadoop Distributed File System (HDFS). They also run the Task Tracker daemon and perform other parallel computation tasks on data that installed applications require (ie. YARN NodeManager daemons, Hadoop MapReduce tasks, and Spark executors). Ideal for reserved/on-demand instances.

### Task
Optional. You can use them to add power to perform parallel computation tasks on data, such as Hadoop MapReduce tasks and Spark executors. Ideal for spot instances. EMR has default functionality for scheduling YARN jobs so that running jobs don’t fail when task nodes running on Spot Instances are terminated by allowing application master processes to run only on core nodes.

## Use Cases

Customer runs an event management SaaS application that uses **Amazon EC2, Auto Scaling, Elastic Load Balancing, and Amazon RDS**. Software is installed on instances at first boot, using Puppet and Chef, which is also used to deploy software updates multiple times per week. A major software overhaul—a new, much larger version of the software—has been deployed to running EC2 instances and some of the instances are being terminated during the update process.

What actions could be taken to prevent instances from being terminated during updates?

- **Use CodeDeploy** to create an application and a deployment targeting the Auto Scaling group. Use CodeDeploy to deploy and update the application.
- **Suspend the Auto Scaling process**. Once suspended, deregister the instance from the ELB, update the application, and register it with the ELB on successful update.

Considerations:
- **Elastic Beanstalk zero downtime feature** cannot update existing applications
- The Auto-scaling process is terminating the instance since it is not responding to health checks and therefore considered unhealthy (and marked for termination).
- Using **termination protection via the console** will **only stop users, not auto-scaling groups** from deleting instances.
- Not a load balancer issue here: detaching a load balancer would stop all traffic to the instances in the ASG.
- **[EC2 autoscaling lifecycle hooks](https://docs.aws.amazon.com/autoscaling/ec2/userguide/lifecycle-hooks.html)**
- If you are using CloudFormation, you can use **[cfn-hup](https://docs.aws.amazon.com/AWSCloudFormation/latest/UserGuide/cfn-hup.html)** (The `cfn-hup` helper is a daemon that detects changes in resource metadata and runs user-specified actions when a change is detected).

## Read-Heavy Workloads

A company has an **application that accesses a relational DB tier** and runs on several EC2 instances placed behind an application load balancer. This application will now be used with a **read-heavy workload** to make medical records accessible to multiple insurance companies. Due to the nature of the data, compliance requires **data protection at-rest and high availability**, especially during sudden large peaks in load.
What is the best solution to improve performance for the current architecture?

**Proposal**: Use an auto scaling group for the EC2 instances across multiple Availability Zones. Deploy an ElastiCache Redis cluster in front of the DB tier and configure **[write through](https://docs.aws.amazon.com/AmazonElastiCache/latest/mem-ug/Strategies.html#Strategies.WriteThrough)**. 

Considerations:
- Memcache does not support encryption at rest for data (must use Redis)
- lazy loading loads data into the cache only when necessary.

## Caching Strategies

### Write Through

#### Advantages

* Data in the cache is never stale. Because the data in the cache is updated every time it's written to the database, the data in the cache is always current.

* Write penalty vs. read penalty. Every write involves two trips:
- A write to the cache
- A write to the database

Which adds latency to the process. That said, **end users are generally more tolerant of latency when updating data than when retrieving data**. There is an inherent sense that updates are more work and thus take longer.

### Lazy Loading
A **cache hit** occurs when data is in the cache and isn't expired:

1. Your application requests data from the cache.

2. The cache returns the data to the application.

A **cache miss** occurs when data isn't in the cache or is expired:

1. Your application requests data from the cache.

2. The cache doesn't have the requested data, so returns a null.

3. Your application requests and receives the data from the database.

4. Your application updates the cache with the new data.

#### Advantages
* Only requested data is cached.

* Because most data is never requested, lazy loading avoids filling up the cache with data that isn't requested.

* Node failures aren't fatal for your application. When a node fails and is replaced by a new, empty node, your application continues to function, though with increased latency. As requests are made to the new node, each cache miss results in a query of the database. At the same time, the data copy is added to the cache so that subsequent requests are retrieved from the cache.

#### Disadvantages

1. Cache misses result in 3 trips:

* Initial request for data from the cache
* Query of the database for the data
* Writing the data to the cache

2. Potential for stale data. 