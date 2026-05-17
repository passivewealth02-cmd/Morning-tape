'use server'

export async function startCheckoutSession(_productId: string): Promise<string | null> {
  return null
}

export async function getCheckoutSessionStatus(_sessionId: string) {
  return { status: null, customerEmail: null }
}
