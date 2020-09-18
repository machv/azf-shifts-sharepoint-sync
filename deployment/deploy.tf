provider "azuread" { 
    version = "~> 0.6"
}
provider "azurerm" { 
    version = "~> 1.35"
}

data "azurerm_client_config" "current" {}

resource "random_string" "appSecret" {
  length = 32
}

resource "azuread_application" "app" {
  name                       = "[TF] Shifts to SharePoint Synchronization"
  homepage                   = "https://localhost"
  reply_urls                 = ["https://localhost:15484/auth"]
  available_to_other_tenants = false
  oauth2_allow_implicit_flow = false
  type                       = "webapp/api"

  required_resource_access {
    resource_app_id = "00000003-0000-0000-c000-000000000000" # Microsoft Graph

    resource_access {
      id   = "e1fe6dd8-ba31-4d61-89e7-88639da4683d" # User.Read
      type = "Scope"
    }

    resource_access {
      id   = "5f8c59db-677d-491f-a6b8-5f174b11ec1d" # Groups.Read.All
      type = "Scope"
    }

    resource_access {
      id   = "89fe6a52-be36-487e-b7d8-d061c450a026" # Sites.ReadWrite.All
      type = "Scope"
    }

    resource_access {
      id   = "a154be20-db9c-4678-8ab7-66f6cc099a59" # User.Read.All
      type = "Scope"
    }
  }
}

resource "azuread_application_password" "app" {
  application_object_id = "${azuread_application.app.object_id}"
  value          = "${random_string.appSecret.result}"
  end_date       = "2030-01-01T00:00:00Z"
}

resource "azurerm_resource_group" "content" {
  name     = "tf-spsync-rg"
  location = "westeurope"
}

resource "azurerm_key_vault" "vault" {
  name                        = "tf-spsync-vault"
  location                    = "${azurerm_resource_group.content.location}"
  resource_group_name         = "${azurerm_resource_group.content.name}"
  tenant_id                   = "${data.azurerm_client_config.current.tenant_id}"

  sku_name = "standard"

  access_policy {
    tenant_id = "${azurerm_function_app.functionApp.identity[0].tenant_id}"
    object_id = "${azurerm_function_app.functionApp.identity[0].principal_id}"

    secret_permissions = [
      "get",
    ]
  }

  access_policy {
    tenant_id = "${data.azurerm_client_config.current.tenant_id}"
    object_id = "${data.azurerm_client_config.current.object_id}"

    secret_permissions = [
        "get",
        "list",
        "set",
        "delete",
        "recover",
        "backup",
        "restore"
    ]
  }

  network_acls {
    default_action = "Allow"
    bypass         = "AzureServices"
  }
}

resource "azurerm_key_vault_secret" "appSecret" {
  name         = "ClientSecret"
  value        = "${random_string.appSecret.result}"
  key_vault_id = "${azurerm_key_vault.vault.id}"
}

resource "azurerm_storage_account" "appStorage" {
  name                     = "tfspsyncstorage"
  resource_group_name      = "${azurerm_resource_group.content.name}"
  location                 = "${azurerm_resource_group.content.location}"
  account_tier             = "Standard"
  account_replication_type = "LRS"
}

resource "azurerm_application_insights" "insights" {
  name                = "tf-spsync-insights"
  location            = "West Europe"
  resource_group_name = "${azurerm_resource_group.content.name}"
  application_type    = "web"
}

resource "azurerm_app_service_plan" "plan" {
  name                = "tf-spsync-plan"
  location            = "${azurerm_resource_group.content.location}"
  resource_group_name = "${azurerm_resource_group.content.name}"
  kind                = "FunctionApp"

  sku {
    tier = "Dynamic"
    size = "Y1"
  }
}

resource "azurerm_function_app" "functionApp" {
  name                      = "tf-spsync-app"
  location                  = "${azurerm_resource_group.content.location}"
  resource_group_name       = "${azurerm_resource_group.content.name}"
  app_service_plan_id       = "${azurerm_app_service_plan.plan.id}"
  storage_connection_string = "${azurerm_storage_account.appStorage.primary_connection_string}"
  app_settings = {
      "FUNCTIONS_WORKER_RUNTIME" = "powershell"
      "FUNCTIONS_EXTENSION_VERSION" = "~2"
      "WEBSITE_CONTENTSHARE" = "spsyncapp"
      "AzureWebJobsStorage" = "${azurerm_storage_account.appStorage.primary_connection_string}"
      "WEBSITE_CONTENTAZUREFILECONNECTIONSTRING" = "${azurerm_storage_account.appStorage.primary_connection_string}"
      "APPINSIGHTS_INSTRUMENTATIONKEY" = "${azurerm_application_insights.insights.instrumentation_key}"
      "APPLICATION_SECRET" = ""
      "REFRESH_TOKEN" = ""
  }
  identity {
      type = "SystemAssigned"
  }
}

output "applicationSecret" {
  value = "${azurerm_key_vault_secret.appSecret.id}"
}

output "applicationId" {
    value = "${azuread_application.app.application_id}"
}
