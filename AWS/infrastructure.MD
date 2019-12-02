# CloudFormation


## Best Practices

### Using Maps
When using EC2 instances across regions, use `Fn:FindInMap` to find the appropriate AMI ID per region, since EC2 instance AMI IDs will differ from region to region.

#### Example:
```yaml
Mappings: 
  RegionMap: 
    us-east-1: 
      HVM64: "ami-0ff8a91507f77f867"
      HVMG2: "ami-0a584ac55a7631c0c"
    us-west-1: 
      HVM64: "ami-0bdb828fd58c52235"
      HVMG2: "ami-066ee5fd4a9ef77f1"
    eu-west-1: 
      HVM64: "ami-047bb4163c506cd98"
      HVMG2: "ami-31c2f645"
Resources: 
  myEC2Instance: 
    Type: "AWS::EC2::Instance"
    Properties: 
      ImageId: !FindInMap
        - RegionMap
        - !Ref 'AWS::Region'
        - HVM64
      InstanceType: m1.small
```
`Fn::FindInMap` for a template with a Mappings section that contains a single map, RegionMap, that associates AMIs with AWS regions. The map has **top-level keys** that correspond to various AWS regions.

Each top-level key is assigned a list with two **second level keys**, `"HVM64"` and `"HVMG2"`, that correspond to the AMI's architecture. Each of the second-level keys is assigned an appropriate AMI name.

The example template contains an `AWS::EC2::Instance` resource whose ImageId property is set by the FindInMap function. MapName is set to the map of interest, "RegionMap" in this example. TopLevelKey is set to the region where the stack is created, which is determined by using the "AWS::Region" pseudo parameter. SecondLevelKey is set to the desired architecture, "HVM64" for this example.

`FindInMap` returns the AMI assigned to FindInMap. For a HVM64 instance in `us-east-1`, `FindInMap` would return `"ami-0ff8a91507f77f867"`.

### Linting
Lint your CloudFormation templates using tools like **[yamllint](https://github.com/adrienverge/yamllint)**, **[cfn-lint](https://github.com/aws-cloudformation/cfn-python-lint)**, or `aws cloudformation validate-template`. Integrate `yamllint` or `cfn-lint` into your editor (like VS Code) and then `aws cloudformation validate-template` prior to `git commit`ing.

### Creation Policies
If you frequently provision EC2 instances with lots of user data, Creation Policies to control instance creation. By default, when CloudFormation creates and EC2 instance it will not wait for the operating system and application to be ready. With a creation policy, you can ask CloudFormation to wait for an external signal.

#### Example:
```yaml
#[...]
Resources:
  VirtualMachine:
    Type: 'AWS::EC2::Instance'
    CreationPolicy:
      ResourceSignal:
        Timeout: PT10M # wait for signal w/ timeout of 10 minutes
    Properties:
      #[...] 
      UserData:
        'Fn::Base64': !Sub |
          #!/bin/bash -ex
          # run some commands
          /opt/aws/bin/cfn-signal --exit-code 0 --resource VirtualMachine --region ${AWS::Region} --stack ${AWS::StackName}
```
The above CreationPolicy that instructs CloudFormation to wait for a signal with a timeout of 10 minutes (PT10M). The `/opt/aws/bin/cfn-signal` command sends the signal to CloudFormation to indicate that the user data script finished successfully (`--exit-code 0`).