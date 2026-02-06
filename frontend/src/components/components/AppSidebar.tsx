import { ChevronDown, FileCog, Files, LandPlot, Settings } from "lucide-react";
import {
  Sidebar,
  SidebarContent,
  SidebarFooter,
  SidebarGroup,
  SidebarGroupContent,
  SidebarHeader,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
} from "../ui/sidebar";
import { Link } from "react-router-dom";
import CustomNavLink from "./CustomNavLink";

export default function AppSidebar() {
  return (
    <Sidebar variant="inset">
      <SidebarHeader>
        <SidebarMenu>
          <SidebarMenuItem>
            <SidebarMenuButton asChild className="">
              <Link to="/">
                <LandPlot />
                <span className="text-base font-semibold">
                  Copro's files Generator
                </span>
              </Link>
            </SidebarMenuButton>
          </SidebarMenuItem>
        </SidebarMenu>
      </SidebarHeader>

      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupContent className="flex flex-col gap-3">
            <SidebarMenu>
              <SidebarMenuItem>
                <CustomNavLink 
                    to="/" 
                    Icon={FileCog}
                    text="Générer les fichiers" />
              </SidebarMenuItem>
            </SidebarMenu>

            <SidebarMenu>
              <SidebarMenuItem>
                 <CustomNavLink 
                    to="/parametres" 
                    Icon={Settings}
                    text="Paramètres" />
              </SidebarMenuItem>
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>
      </SidebarContent>
      <SidebarFooter />
    </Sidebar>
  );
}
