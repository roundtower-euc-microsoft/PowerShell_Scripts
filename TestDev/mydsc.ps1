Configuration MyDSCConfiguration {
    Import-DSCResource -Module NetworkingDSC
    Node "localhost" {
        DnsServerAddress DnsServerAddress
            {
            address = '1.1.1.1','1.0.0.1'
            InterfaceAlias = 'Ethernet'
            AddressFamily = 'IPv4'
            Validate = $true
            }
        FireWall BlockWeb 
            {
            Action = 'Block'
            Name = 'BlockWeb'
            DisplayName = 'Block Web Access'
            Group = 'Block Web Rule Group'
            Ensure = 'Present'
            Enabled = 'True'
            Profile = ('Domain', 'Private','Public')
            Direction = 'OutBound'
            RemotePort = ('80', '443')
            Protocol = 'TCP'
            Description = 'Rule to prevent access to most web sites'
            }
        FirewallProfile ConfigurePublicFW 
            {
            Name = 'Public'
            Enabled = 'True'
            DefaultInboundAction = 'Allow'
            DefaultOutboundAction	= 'Allow'
            AllowInboundRules = 'True'
            AllowLocalFirewallRules = 'True'
            }
        FirewallProfile ConfigurePrivateFW 
            {
            Name = 'Private'
            Enabled = 'True'
            DefaultInboundAction = 'Allow'
            DefaultOutboundAction	= 'Allow'
            AllowInboundRules = 'True'
            AllowLocalFirewallRules = 'True'
            }
        FirewallProfile ConfigureDomainFW 
            {
            Name = 'Domain'
            Enabled = 'True'
            DefaultInboundAction = 'Allow'
            DefaultOutboundAction	= 'Allow'
            AllowInboundRules = 'True'
            AllowLocalFirewallRules = 'True'
            }
        }
    }