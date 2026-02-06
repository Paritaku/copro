import { Separator } from "../ui/separator";
import { SidebarTrigger } from "../ui/sidebar";

export default function CustomBeadCrumb({
    pageTitle="Page Title",
    children

} : {
    pageTitle?: string,
    children: React.ReactElement
}){
    return (
         <main className="py-2 px-4">
            <header>
                <div className="flex items-center min-h-12 gap-3">
                    <SidebarTrigger/>
                    <h1 className="text-lg font-semibold leading-tight text-foreground sm:text-xl self-center">{pageTitle}</h1>
                </div>
                <Separator className="mt-1"/>
            </header>
            <div className="mt-2">
                {children}
            </div>
        </main>
    )
}