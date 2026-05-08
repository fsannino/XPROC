'use client'

import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { useRouter } from 'next/navigation'
import {
  ReactFlow,
  ReactFlowProvider,
  Background,
  Controls,
  MiniMap,
  type Node,
  type Edge,
  type NodeChange,
  applyNodeChanges,
  type OnNodesChange,
} from '@xyflow/react'
import '@xyflow/react/dist/style.css'
import type { MapaNode, MapaEdge } from '@/actions/mapa'
import { upsertNode, updateNodePosition, deleteNode } from '@/actions/mapa'
import { MAPA_LEVELS, type NodeType } from '@/lib/definitions'
import NodeDrawer, { type DrawerState } from '@/components/mapa/NodeDrawer'
import CustomNode from '@/components/mapa/CustomNode'

type Props = {
  initialNodes: MapaNode[]
  initialEdges: MapaEdge[]
}

const nodeTypes = { xproc: CustomNode }

/** Layout inicial automático para nós em (0,0): grade por tipo. */
function autoLayout(node: MapaNode, index: number): { x: number; y: number } {
  if (node.posicaoX !== 0 || node.posicaoY !== 0) {
    return { x: node.posicaoX, y: node.posicaoY }
  }
  const yByType: Record<NodeType, number> = {
    cadeia: 0, macroprocesso: 180, processo: 360,
    macroatividade: 540, atividade: 720,
  }
  return { x: 60 + index * 260, y: yByType[node.tipo] }
}

function nodeKey(tipo: NodeType, id: number) {
  return `${tipo}:${id}`
}

function parseKey(key: string): { tipo: NodeType; id: number } {
  const [tipo, id] = key.split(':') as [NodeType, string]
  return { tipo, id: Number(id) }
}

export default function MapaCanvas(props: Props) {
  return (
    <ReactFlowProvider>
      <CanvasInner {...props} />
    </ReactFlowProvider>
  )
}

function CanvasInner({ initialNodes, initialEdges }: Props) {
  const router = useRouter()

  const grouped = useMemo(() => {
    const indexByType: Record<NodeType, number> = {
      cadeia: 0, macroprocesso: 0, processo: 0, macroatividade: 0, atividade: 0,
    }
    return initialNodes.map((n) => {
      const idx = indexByType[n.tipo]++
      const pos = autoLayout(n, idx)
      const node: Node = {
        id: nodeKey(n.tipo, n.id),
        type: 'xproc',
        position: pos,
        data: {
          tipo: n.tipo,
          label: n.descricao,
          abreviacao: n.abreviacao,
          sequencia: n.sequencia,
          onAddChild: () => openDrawerFor({ mode: 'create', tipoFilho: childTipo(n.tipo), parentId: n.id }),
          onDelete: () => handleDelete(n.tipo, n.id),
        },
      }
      return node
    })
  }, [initialNodes]) // eslint-disable-line react-hooks/exhaustive-deps

  const initialFlowEdges: Edge[] = useMemo(
    () =>
      initialEdges.map((e) => ({
        id: `e:${e.source}->${e.target}`,
        source: e.source,
        target: e.target,
        type: 'smoothstep',
        style: { stroke: '#1A6E8E', strokeWidth: 1.5 },
      })),
    [initialEdges],
  )

  const [nodes, setNodes] = useState<Node[]>(grouped)
  const [edges] = useState<Edge[]>(initialFlowEdges)
  const [drawer, setDrawer] = useState<DrawerState | null>(null)
  const dragTimers = useRef(new Map<string, ReturnType<typeof setTimeout>>())

  const openDrawerFor = useCallback((s: DrawerState) => setDrawer(s), [])

  const handleDelete = useCallback(async (tipo: NodeType, id: number) => {
    if (!confirm(`Excluir ${MAPA_LEVELS[tipo].label.toLowerCase()}?`)) return
    const res = await deleteNode({ tipo, id })
    if (!res.success) {
      alert(res.error ?? 'Falha ao excluir.')
      return
    }
    setNodes((ns) => ns.filter((n) => n.id !== nodeKey(tipo, id)))
    router.refresh()
  }, [router])

  const onNodesChange: OnNodesChange = useCallback((changes: NodeChange[]) => {
    setNodes((ns) => applyNodeChanges(changes, ns))
    // Persistir posições com debounce por nó
    for (const ch of changes) {
      if (ch.type !== 'position' || !ch.position || ch.dragging) continue
      const { tipo, id } = parseKey(ch.id)
      const prev = dragTimers.current.get(ch.id)
      if (prev) clearTimeout(prev)
      const t = setTimeout(() => {
        updateNodePosition({
          tipo,
          id,
          posicaoX: ch.position!.x,
          posicaoY: ch.position!.y,
        }).catch(() => undefined)
      }, 250)
      dragTimers.current.set(ch.id, t)
    }
  }, [])

  const onNodeClick = useCallback((_: React.MouseEvent, n: Node) => {
    const { tipo, id } = parseKey(n.id)
    const data = n.data as { label: string; abreviacao?: string | null; sequencia?: number | null }
    setDrawer({
      mode: 'edit',
      tipo,
      id,
      descricao: data.label,
      abreviacao: data.abreviacao ?? '',
      sequencia: data.sequencia ?? undefined,
    })
  }, [])

  const handleSubmit = useCallback(async (payload: {
    mode: 'create' | 'edit'
    tipo: NodeType
    id?: number
    parentId?: number
    descricao: string
    abreviacao?: string
    sequencia?: number
  }) => {
    const res = await upsertNode(payload)
    if (!res?.success) {
      alert(res?.error ?? 'Falha ao salvar.')
      return false
    }
    setDrawer(null)
    // router.refresh() invalida o Router Cache do Next, então quando o
    // usuário navegar para /processos, /transacoes, etc., os dados serão
    // re-buscados do banco (com revalidatePath em layout no server action).
    router.refresh()
    return true
  }, [router])

  // Fechar timers ao desmontar
  useEffect(() => {
    const map = dragTimers.current
    return () => map.forEach((t) => clearTimeout(t))
  }, [])

  return (
    <div className="relative w-full h-full">
      <ReactFlow
        nodes={nodes}
        edges={edges}
        nodeTypes={nodeTypes}
        onNodesChange={onNodesChange}
        onNodeClick={onNodeClick}
        fitView
        proOptions={{ hideAttribution: true }}
      >
        <Background color="#E2E8F0" gap={20} />
        <Controls position="bottom-right" showInteractive={false} />
        <MiniMap
          pannable zoomable
          nodeColor={() => '#1A6E8E'}
          maskColor="rgba(11,61,92,0.06)"
        />
      </ReactFlow>

      <button
        onClick={() => setDrawer({ mode: 'create', tipoFilho: 'cadeia' })}
        className="absolute top-4 left-4 bg-navy hover:bg-teal text-white text-sm font-semibold px-4 py-2 rounded-md shadow-md transition-all hover:-translate-y-0.5"
      >
        + Cadeia de Valor
      </button>

      {drawer && (
        <NodeDrawer
          state={drawer}
          onClose={() => setDrawer(null)}
          onSubmit={handleSubmit}
        />
      )}
    </div>
  )
}

function childTipo(tipo: NodeType): NodeType | null {
  const order: NodeType[] = ['cadeia', 'macroprocesso', 'processo', 'macroatividade', 'atividade']
  const idx = order.indexOf(tipo)
  return idx >= 0 && idx < order.length - 1 ? order[idx + 1] : null
}
