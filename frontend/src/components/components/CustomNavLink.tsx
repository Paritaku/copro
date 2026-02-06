import { NavLink } from "react-router-dom";
import { SidebarMenuButton } from "../ui/sidebar";
import type { LucideProps } from "lucide-react";

export default function CustomNavLink({
    to,
    activeClass = "bg-primary text-primary-foreground hover:bg-primary/80 hover:text-primary-foreground",
    Icon,
    text,
} : {
    to: string,
    activeClass?: string,
    Icon: React.ForwardRefExoticComponent<Omit<LucideProps, "ref"> & React.RefAttributes<SVGSVGElement>>
    text: string

}){
    return (
        <NavLink to={to}>
            {({isActive}) => (
                <SidebarMenuButton className={isActive ? activeClass : ""}>
                      <Icon />
                      <span>{text}</span>
                </SidebarMenuButton>
            )}
        </NavLink>
    )
}