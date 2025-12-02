import { Request, Response, NextFunction } from 'express';
export interface AuthRequest extends Request {
    userId?: number;
}
export declare function authenticateToken(req: AuthRequest, res: Response, next: NextFunction): Response<any, Record<string, any>> | undefined;
export declare function authenticateTokenVercel(req: any): Promise<{
    success: boolean;
    userId?: number;
    message?: string;
}>;
export declare function generateToken(userId: number): string;
//# sourceMappingURL=auth.d.ts.map