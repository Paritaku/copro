import { SidebarInset, SidebarProvider, SidebarTrigger } from "../ui/sidebar"
import { Outlet } from "react-router-dom"
import AppSidebar from "./AppSidebar"
import { Separator } from "../ui/separator"

export default function Layout() {
    return (
        <SidebarProvider>
            <AppSidebar />
            <SidebarInset>
                <Outlet />
            </SidebarInset>
        </SidebarProvider>
    )
}